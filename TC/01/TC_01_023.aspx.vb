Partial Class TC_01_023
    Inherits AuthBasePage

    Const cst_rqProcessType As String = "ProcessType"
    Const cst_rqProcessType_Update As String = "Update"
    Const cst_rqProcessType_Insert As String = "Insert"
    Const cst_rqProcessType_View As String = "View"

    Const cst_cmd_edit As String = "edit"
    Const cst_cmd_view As String = "view"
    Const cst_cmd_return As String = "return"
    Const cst_cmd_stop As String = "stop"
    Const cst_cmd_del As String = "del"
    'Dim lrMsg As String = ""

    Const Cst_SearchName As String = "_searchTC_01_023"
    'Const Upload_Path As String = "~/images/Placepic/"
    'Dim strSyncUploadEnable As String = System.Configuration.ConfigurationSettings.AppSettings("syncUploadEnable") '是否同步
    'Dim strSyncUploadServer As String = System.Configuration.ConfigurationSettings.AppSettings("syncUploadServer")  '目標Server
    'Dim strSyncUploadPort As String = System.Configuration.ConfigurationSettings.AppSettings("syncUploadPort")  '目標Server
    'Dim strSyncUploadPW As String = System.Configuration.ConfigurationSettings.AppSettings("syncUploadPW")  '同步目標Server的密碼
    'Dim strSyncUploadUser As String = System.Configuration.ConfigurationSettings.AppSettings("syncUploadUser")  '同步目標Server的帳號
    'Dim strSyncUploadFolder As String = "/upload/Placepic/"  '附件路徑

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not Me.IsPostBack Then
            Call cCreate1()
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
    End Sub

    Sub cCreate1()
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        orgid_value.Value = TIMS.Get_OrgID(RIDValue.Value, objconn)
        If orgid_value.Value = "" OrElse orgid_value.Value = "-1" Then orgid_value.Value = sm.UserInfo.OrgID

        DataGridTable.Visible = False
        If Session(Cst_SearchName) IsNot Nothing Then
            Call UseKeepSearch()
            Return
        End If

        '第1次進入就查詢
        Call sUtl_Search1()
        'bt_search_Click(sender, e)
    End Sub

    Sub UseKeepSearch()
        If Session(Cst_SearchName) = Nothing Then Return
        Dim MyValue As String = ""
        Dim str_Search1 As String = Session(Cst_SearchName)
        Session(Cst_SearchName) = Nothing
        If str_Search1 = "" Then Return

        MyValue = TIMS.GetMyValue(str_Search1, "prgid")
        If MyValue = "" OrElse MyValue <> TIMS.Get_MRqID(Me) Then Return

        center.Text = TIMS.GetMyValue(str_Search1, "center")
        RIDValue.Value = TIMS.GetMyValue(str_Search1, "RIDValue")
        orgid_value.Value = TIMS.GetMyValue(str_Search1, "orgid_value")

        RMTNO.Text = TIMS.GetMyValue(str_Search1, "RMTNO")
        RMTNAME.Text = TIMS.GetMyValue(str_Search1, "RMTNAME")

        PageControler1.PageIndex = 0
        MyValue = TIMS.GetMyValue(str_Search1, "PageIndex")
        If MyValue <> "" AndAlso IsNumeric(MyValue) Then
            MyValue = CInt(MyValue)
            PageControler1.PageIndex = MyValue
        End If
        MyValue = TIMS.GetMyValue(str_Search1, "submit")
        If MyValue = "1" Then
            '有記憶值進入就查詢
            Call sUtl_Search1()
            Return
        End If

    End Sub

    Sub KeepSearch()
        Session(Cst_SearchName) = Nothing

        'Dim s_submit As String = If(DataGridTable.Visible, "1", "0")
        Dim str_Search1 As String = ""
        str_Search1 &= "&prgid=" & TIMS.ClearSQM(Request("ID"))
        str_Search1 &= "&center=" & TIMS.ClearSQM(center.Text)
        str_Search1 &= "&RIDValue=" & TIMS.ClearSQM(RIDValue.Value)
        str_Search1 &= "&orgid_value=" & TIMS.ClearSQM(orgid_value.Value)
        str_Search1 &= "&RMTNO=" & TIMS.ClearSQM(RMTNO.Text)
        str_Search1 &= "&RMTNAME=" & TIMS.ClearSQM(RMTNAME.Text)
        str_Search1 &= "&PageIndex=" & DataGrid1.CurrentPageIndex + 1
        str_Search1 &= "&submit=" & If(DataGridTable.Visible, "1", "0")

        Session(Cst_SearchName) = str_Search1
    End Sub

    '查詢 search
    Sub sUtl_Search1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        msg.Text = "查無資料!!"
        DataGridTable.Visible = False

        RMTNO.Text = TIMS.ClearSQM(RMTNO.Text)
        RMTNAME.Text = TIMS.ClearSQM(RMTNAME.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        Dim hPMS As New Hashtable
        Dim sSql As String = ""
        sSql &= " SELECT a.RMTID ,a.RMTNO,a.RMTNAME ,a.ORGID" & vbCrLf
        sSql &= " ,a.KTSID,dbo.FN_TEACHSOFT_N(a.KTSID) TEACHSOFT_N" & vbCrLf
        sSql &= " ,a.KTDID,dbo.FN_TEACHDEVICE_N(a.KTDID) TEACHDEVICE_N" & vbCrLf
        sSql &= " ,a.CABLENETWORK,a.CABLEDLRATE,a.CABLEUPRATE" & vbCrLf
        sSql &= " ,a.WIFINETWORK,a.WIFIDLRATE,a.WIFIUPRATE" & vbCrLf
        sSql &= " ,a.VIDEODEVICE" & vbCrLf
        sSql &= " ,a.SOFTDESC,a.DEVICEDESC" & vbCrLf
        sSql &= " ,a.RMTPIC1,a.RMTPIC2,a.RMTPIC3,a.RMTPIC4" & vbCrLf
        sSql &= " ,a.MODIFYACCT,a.MODIFYDATE,a.MODIFYTYPE" & vbCrLf
        sSql &= " FROM ORG_REMOTER a" & vbCrLf
        sSql &= " JOIN ORG_ORGINFO oo on oo.ORGID=a.ORGID" & vbCrLf
        sSql &= " WHERE 1=1" & vbCrLf
        '含已停用資料 MODIFYTYPE:D
        If Not chkboxDelData.Checked Then sSql &= " AND a.MODIFYTYPE IS NULL" & vbCrLf

        If RIDValue.Value <> "" Then
            If sm.UserInfo.LID = 0 AndAlso RIDValue.Value.Length = 1 Then
                Dim s_DISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
                hPMS.Add("DISTID", s_DISTID)
                hPMS.Add("YEARS", CStr(sm.UserInfo.Years))
                sSql &= " AND a.ORGID IN (SELECT ORGID FROM VIEW_RIDNAME WHERE DISTID=@DISTID AND YEARS=@YEARS)" & vbCrLf
            Else
                hPMS.Add("RID", RIDValue.Value)
                sSql &= " AND a.ORGID IN (SELECT ORGID FROM AUTH_RELSHIP WHERE RID=@RID)" & vbCrLf
            End If
        Else
            hPMS.Add("ORGID", sm.UserInfo.OrgID)
            sSql &= " AND a.ORGID=@ORGID" & vbCrLf
        End If
        If RMTNO.Text <> "" Then
            Dim lk_RMTNO As String = String.Concat("%", RMTNO.Text, "%")
            hPMS.Add("lk_RMTNO", lk_RMTNO)
            sSql &= " AND a.RMTNO LIKE @lk_RMTNO" & vbCrLf
        End If
        If RMTNAME.Text <> "" Then
            Dim lk_RMTNAME As String = String.Concat("%", RMTNAME.Text, "%")
            hPMS.Add("lk_RMTNAME", lk_RMTNAME)
            sSql &= " AND a.RMTNAME LIKE @lk_RMTNAME" & vbCrLf
        End If
        sSql &= " ORDER BY a.MODIFYDATE DESC"
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, hPMS)

        msg.Text = "查無資料!!"
        DataGridTable.Visible = False
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        msg.Text = ""
        DataGridTable.Visible = True

        PageControler1.PageDataTable = dt
        PageControler1.PrimaryKey = "RMTID"
        'PageControler1.Sort = "PlaceID"
        PageControler1.Sort = "MODIFYDATE DESC, RMTNO ASC"  'edit，by:20181026
        PageControler1.ControlerLoad()
    End Sub


    ''' <summary>'新增</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnAddNew1_Click(sender As Object, e As EventArgs) Handles BtnAddNew1.Click
        RMTNO.Text = TIMS.ClearSQM(RMTNO.Text)
        RMTNAME.Text = TIMS.ClearSQM(RMTNAME.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        orgid_value.Value = TIMS.Get_OrgID(RIDValue.Value, objconn)
        If orgid_value.Value = "" OrElse orgid_value.Value = "-1" Then orgid_value.Value = sm.UserInfo.OrgID

        Dim RqID1 As String = TIMS.Get_MRqID(Me)

        Call KeepSearch()
        'Response.Redirect("TC_01_013_add.aspx?ProcessType=Insert&ID=" & Request("ID") & "")
        '20100208 按新增時代查詢之 場地代碼 & 場地名稱
        'Response.Redirect("TC_01_013_add.aspx?ProcessType=Insert&ID=" & Request("ID") & "&PlaceNo=" & PlaceNo.Text & "&Place=" & Place.Text & "&RID=" & RIDValue.Value)
        Dim s_MyValue1 As String = ""
        TIMS.SetMyValue(s_MyValue1, "RMTNO", RMTNO.Text)
        TIMS.SetMyValue(s_MyValue1, "RMTNAME", RMTNAME.Text)
        TIMS.SetMyValue(s_MyValue1, "ProcessType", cst_rqProcessType_Insert)
        TIMS.SetMyValue(s_MyValue1, "RID", RIDValue.Value)
        TIMS.SetMyValue(s_MyValue1, "OrgID", orgid_value.Value)
        Dim url1 As String = String.Concat("TC_01_023_ADD?ID=", RqID1, s_MyValue1)
        Call TIMS.Utl_Redirect(Me, objconn, url1)
        'sm.RedirectUrlAfterBlock = ResolveUrl(url1)
    End Sub

    ''' <summary>查詢</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnSearch1_Click(sender As Object, e As EventArgs) Handles BtnSearch1.Click
        Call sUtl_Search1()
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim lbtView As LinkButton = e.Item.FindControl("lbtView")
                Dim lbtEdit As LinkButton = e.Item.FindControl("lbtEdit")
                Dim lbtReturn As LinkButton = e.Item.FindControl("lbtReturn")
                Dim lbtStop As LinkButton = e.Item.FindControl("lbtStop")
                Dim lbtDel As LinkButton = e.Item.FindControl("lbtDel")
                e.Item.Cells(0).Text = DataGrid1.CurrentPageIndex * DataGrid1.PageSize + e.Item.ItemIndex + 1
                lbtReturn.Attributes("onclick") = "return confirm('確定要啟用?');"
                lbtStop.Attributes("onclick") = "return confirm('確定要停用?');"
                lbtDel.Attributes("onclick") = "return confirm('確定要刪除?');"

                Dim cmdArg As String = ""
                TIMS.SetMyValue(cmdArg, "RMTID", drv("RMTID"))
                TIMS.SetMyValue(cmdArg, "ORGID", drv("ORGID"))
                lbtView.CommandArgument = cmdArg 'Convert.ToString(drv("PTID"))
                lbtEdit.CommandArgument = cmdArg 'Convert.ToString(drv("PTID"))
                lbtDel.CommandArgument = cmdArg 'Convert.ToString(drv("PTID"))
                lbtReturn.CommandArgument = cmdArg '
                lbtStop.CommandArgument = cmdArg 'Convert.ToString(drv("PTID"))

                Dim fg_CAN_Return As Boolean = (Convert.ToString(drv("MODIFYTYPE")).ToUpper() = "D")
                '分署／可刪、修、檢視、還原
                lbtView.Style("display") = If(fg_CAN_Return, "none", "")  '可檢視／'不能檢視
                lbtEdit.Style("display") = If(fg_CAN_Return, "none", "")  '可編輯／'不能編輯
                lbtStop.Style("display") = If(fg_CAN_Return, "none", "")  '可停用／'不能停用
                lbtReturn.Style("display") = If(fg_CAN_Return, "", "none")  '(啟用)有還原／'不能還原
                lbtDel.Style("display") = "" '永遠有刪
                'lbtDel.Style("display") = If(fg_CAN_Return, "", "none")  '有刪／'不能刪

                '有停用 就沒有還原
                lbtView.Visible = False
                'lbtReturn.Visible = If(fg_CAN_Return, True, False)
                'lbtStop.Visible = If(fg_CAN_Return, False, True)
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e Is Nothing OrElse e.CommandArgument = "" Then Return

        Dim s_CmdArg As String = e.CommandArgument
        Dim sRMTID As String = TIMS.GetMyValue(s_CmdArg, "RMTID")
        Dim sORGID As String = TIMS.GetMyValue(s_CmdArg, "ORGID")
        Dim RqID1 As String = TIMS.Get_MRqID(Me)
        If sRMTID = "" OrElse sORGID = "" OrElse RIDValue.Value = "" Then Return
        orgid_value.Value = sORGID

        Dim lrMsg As String = ""
        Select Case e.CommandName
            Case cst_cmd_view '檢視-for 訓練單位
                KeepSearch()
                Dim s_MyValue1 As String = ""
                TIMS.SetMyValue(s_MyValue1, "RMTID", sRMTID)
                TIMS.SetMyValue(s_MyValue1, cst_rqProcessType, cst_rqProcessType_View)
                TIMS.SetMyValue(s_MyValue1, "RID", RIDValue.Value)
                TIMS.SetMyValue(s_MyValue1, "OrgID", orgid_value.Value)
                Dim url1 As String = String.Concat("TC_01_023_ADD?ID=", RqID1, s_MyValue1)
                Call TIMS.Utl_Redirect(Me, objconn, url1)

            Case cst_cmd_edit
                KeepSearch()
                Dim s_MyValue1 As String = ""
                TIMS.SetMyValue(s_MyValue1, "RMTID", sRMTID)
                TIMS.SetMyValue(s_MyValue1, cst_rqProcessType, cst_rqProcessType_Update)
                TIMS.SetMyValue(s_MyValue1, "RID", RIDValue.Value)
                TIMS.SetMyValue(s_MyValue1, "OrgID", orgid_value.Value)
                Dim url1 As String = String.Concat("TC_01_023_ADD?ID=", RqID1, s_MyValue1)
                Call TIMS.Utl_Redirect(Me, objconn, url1)

            Case cst_cmd_return '還原
                Dim rfg_OK1 As Boolean = UPDATE_REMOTER_MODIFYTYPE(sRMTID, cst_cmd_return)
                lrMsg = If(rfg_OK1, "還原成功!!", "還原失敗!!")
                sm.LastResultMessage = lrMsg
                Call sUtl_Search1() '還原成功 可查詢

            Case cst_cmd_stop '停用
                Dim rfg_OK1 As Boolean = UPDATE_REMOTER_MODIFYTYPE(sRMTID, cst_cmd_stop)
                lrMsg = If(rfg_OK1, "停用成功!!", "停用失敗!!")
                sm.LastResultMessage = lrMsg
                Call sUtl_Search1() '停用成功 可查詢

            Case cst_cmd_del '刪除
                If ChkRMTIDUse(sRMTID, objconn) Then
                    lrMsg = "遠距課程環境目前被使用，不可刪除!!"
                    sm.LastResultMessage = lrMsg
                    Exit Sub
                End If
                Dim rfg_OK1 As Boolean = UPDATE_REMOTER_MODIFYTYPE(sRMTID, cst_cmd_del)
                lrMsg = If(rfg_OK1, "刪除成功!!", "刪除失敗!!")
                sm.LastResultMessage = lrMsg
                Call sUtl_Search1() '刪除成功 可查詢

        End Select

    End Sub

    Private Function ChkRMTIDUse(sRMTID As String, oConn As SqlConnection) As Boolean
        Dim rst As Boolean = False
        'true: Common.MessageBox(Me, "該場地目前被使用，不可刪除!!")
        sRMTID = TIMS.ClearSQM(sRMTID)
        If sRMTID = "" Then Return rst

        TIMS.OpenDbConn(objconn)
        Dim sql As String = "SELECT 1 FROM PLAN_PLANINFO WHERE RMTID=@RMTID"
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, oConn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("RMTID", SqlDbType.Int).Value = Val(sRMTID)
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then rst = True
        If (rst) Then Return rst
        Return rst 'false(未查詢到使用該遠距課程環境)
    End Function

    Private Function UPDATE_REMOTER_MODIFYTYPE(ByVal sRMTID As String, str_CMD1 As String) As Boolean
        Dim rst As Boolean = False
        Dim iRST As Integer = 0
        sRMTID = TIMS.ClearSQM(sRMTID)
        str_CMD1 = TIMS.ClearSQM(str_CMD1)
        orgid_value.Value = TIMS.ClearSQM(orgid_value.Value)
        If orgid_value.Value = "" OrElse sRMTID = "" OrElse str_CMD1 = "" Then Return rst

        Select Case str_CMD1
            Case cst_cmd_return '還原
                Dim hPMS As New Hashtable From {{"MODIFYTYPE", sm.UserInfo.UserID}, {"RMTID", Val(sRMTID)}, {"ORGID", Val(orgid_value.Value)}}
                Dim usSql As String = ""
                usSql &= " UPDATE ORG_REMOTER" & vbCrLf
                usSql &= " SET MODIFYTYPE=NULL, MODIFYACCT=@MODIFYTYPE, MODIFYDATE=GETDATE()" & vbCrLf
                usSql &= " WHERE RMTID=@RMTID AND ORGID=@ORGID" & vbCrLf
                iRST = DbAccess.ExecuteNonQuery(usSql, objconn, hPMS)
                rst = (iRST > 0)

            Case cst_cmd_stop '停用 '停用／'刪除
                Dim hPMS As New Hashtable From {{"MODIFYTYPE", sm.UserInfo.UserID}, {"RMTID", Val(sRMTID)}, {"ORGID", Val(orgid_value.Value)}}
                Dim usSql As String = ""
                usSql &= " UPDATE ORG_REMOTER" & vbCrLf
                usSql &= " SET MODIFYTYPE='D', MODIFYACCT=@MODIFYTYPE, MODIFYDATE=GETDATE()" & vbCrLf
                usSql &= " WHERE RMTID=@RMTID AND ORGID=@ORGID" & vbCrLf
                iRST = DbAccess.ExecuteNonQuery(usSql, objconn, hPMS)
                rst = (iRST > 0)

            Case cst_cmd_del '刪除 '停用／'刪除
                'Dim usSql As String = ""
                'usSql &= " UPDATE ORG_REMOTER" & vbCrLf
                'usSql &= " SET MODIFYTYPE='D', MODIFYACCT=@MODIFYTYPE, MODIFYDATE=GETDATE()" & vbCrLf
                'usSql &= " WHERE RMTID=@RMTID AND ORGID=@ORGID" & vbCrLf
                If ChkRMTIDUse(sRMTID, objconn) Then Return rst

                'Hid_RMTID.Value = TIMS.ClearSQM(Hid_RMTID.Value)
                orgid_value.Value = TIMS.ClearSQM(orgid_value.Value)

                If sRMTID = "" OrElse orgid_value.Value = "" Then Return rst

                Dim rPMS As New Hashtable From {{"RMTID", Val(sRMTID)}, {"ORGID", Val(orgid_value.Value)}}

                Dim dtPIC As DataTable = TIMS.CreatePICRMTdt(Server, rPMS, objconn)

                Dim Upload_Path As String = TIMS.GET_UPLOADPATH1_RMT()

                Try
                    If dtPIC IsNot Nothing AndAlso dtPIC.Rows.Count > 0 Then
                        For Each dr1 As DataRow In dtPIC.Rows
                            Dim vFILENAME1 As String = Convert.ToString(dr1("FileName1"))
                            Dim flag_PIC_EXISTS As Boolean = TIMS.CHK_PIC_EXISTS(Server, Upload_Path, vFILENAME1)
                            If flag_PIC_EXISTS Then IO.File.Delete(Server.MapPath(String.Concat(Upload_Path, vFILENAME1)))
                        Next
                    End If
                Catch ex As Exception
                    TIMS.LOG.Warn(ex.Message, ex) 'Common.MessageBox(Me, ex.ToString)
                    Dim strErrmsg As String = String.Concat("ex.ToString:", ex.ToString, vbCrLf)
                    TIMS.WriteTraceLog(Me, ex, strErrmsg)
                    'Exit Sub 'Throw ex
                End Try

                Dim hPMS As New Hashtable From {{"MODIFYTYPE", sm.UserInfo.UserID}, {"RMTID", Val(sRMTID)}, {"ORGID", Val(orgid_value.Value)}}
                Dim dsSql As String = ""
                dsSql &= " DELETE ORG_REMOTER" & vbCrLf
                dsSql &= " WHERE RMTID=@RMTID AND ORGID=@ORGID" & vbCrLf
                iRST = DbAccess.ExecuteNonQuery(dsSql, objconn, hPMS)
                rst = True '(iRST > 0)

        End Select
        Return rst
    End Function


End Class