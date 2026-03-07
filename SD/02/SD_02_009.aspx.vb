Partial Class SD_02_009
    Inherits AuthBasePage

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
        '設定分頁
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            DataGridTable.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button12_Click(sender, e)
            End If

        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

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

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確

        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim sSql As String = ""
        sSql &= " SELECT a.PlanID,a.ComIDNO,a.SeqNO,a.OrgName" & vbCrLf
        sSql &= " ,a.ClassCName2 ClassName" & vbCrLf
        sSql &= " ,a.STDate,a.FTDate,a.OCID" & vbCrLf
        sSql &= " FROM VIEW2 a" & vbCrLf
        sSql &= " WHERE a.RID IN (SELECT RID FROM Auth_Relship WHERE RelShip like '" & RelShip & "%')"
        If sm.UserInfo.RID = "A" Then
            sSql &= " and a.TPLANID='" & sm.UserInfo.TPlanID & "' and a.YEARS='" & sm.UserInfo.Years & "'"
        Else
            sSql &= " and a.PLANID='" & sm.UserInfo.PlanID & "'"
        End If
        If OCIDValue1.Value <> "" Then sSql += " and a.OCID = '" & OCIDValue1.Value & "'" & vbCrLf
        If cjobValue.Value <> "" Then sSql += " and a.CJOB_UNKEY='" & cjobValue.Value & "'" & vbCrLf

        DataGridTable.Visible = False
        msg.Text = "查無資料"

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn)
        If TIMS.dtNODATA(dt) Then Return

        DataGridTable.Visible = True
        msg.Text = ""

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "SD_TD1"

            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "SD_TD2"
                End If

                Dim drv As DataRowView = e.Item.DataItem
                Dim btn As HtmlInputButton = e.Item.FindControl("Button3")
                Dim btn1 As HtmlInputButton = e.Item.FindControl("Button4")
                Dim OCID As HtmlInputHidden = e.Item.FindControl("OCID")
                Dim consignee2 As HtmlInputHidden = e.Item.FindControl("consignee2")

                OCID.Value = drv("OCID")

                If consignee.SelectedValue = "01" Then '如果是依錄取通知對象列印,則只印出正取的人選,正取的代碼是01
                    consignee2.Value = "01"

                    btn.Attributes("onclick") = "openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_02_009&OCID='+escape('" & OCID.Value & "')+'&SelResultID='+escape('" & consignee2.Value & "'));"
                    btn1.Attributes("onclick") = "openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_02_009_1&OCID='+escape('" & OCID.Value & "')+'&SelResultID='+escape('" & consignee2.Value & "'));"
                Else
                    btn.Attributes("onclick") = "openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_02_009&OCID='+escape('" & OCID.Value & "')+' &');"
                    btn1.Attributes("onclick") = "openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_02_009_1&OCID='+escape('" & OCID.Value & "')+' &');"
                End If

        End Select


    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub

    Private Sub consignee_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles consignee.SelectedIndexChanged
        DataGridTable.Visible = False
    End Sub
End Class
