Partial Class TC_01_013
    Inherits AuthBasePage

    Const cst_cmd_edit As String = "edit"
    Const cst_cmd_view As String = "view"
    Const cst_cmd_return As String = "return"
    Const cst_cmd_stop As String = "stop"
    Const cst_cmd_del As String = "del"
    Dim lrMsg As String = ""

    Const Cst_SearchName As String = "_search"
    Const Upload_Path As String = "~/images/Placepic/"
    Dim strSyncUploadEnable As String = System.Configuration.ConfigurationSettings.AppSettings("syncUploadEnable") '是否同步
    Dim strSyncUploadServer As String = System.Configuration.ConfigurationSettings.AppSettings("syncUploadServer")  '目標Server
    Dim strSyncUploadPort As String = System.Configuration.ConfigurationSettings.AppSettings("syncUploadPort")  '目標Server
    Dim strSyncUploadPW As String = System.Configuration.ConfigurationSettings.AppSettings("syncUploadPW")  '同步目標Server的密碼
    Dim strSyncUploadUser As String = System.Configuration.ConfigurationSettings.AppSettings("syncUploadUser")  '同步目標Server的帳號
    Dim strSyncUploadFolder As String = "/upload/Placepic/"  '附件路徑

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        '分頁設定 Start
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
            Dim MyValue As String = ""
            MyValue = TIMS.GetMyValue(Session(Cst_SearchName), "prgid")
            If MyValue <> "" AndAlso MyValue = TIMS.Get_MRqID(Me) Then
                'If MyValue = TIMS.Get_MRqID(Me) Then Session(Cst_SearchName) = Nothing
                center.Text = TIMS.GetMyValue(Session(Cst_SearchName), "center")
                RIDValue.Value = TIMS.GetMyValue(Session(Cst_SearchName), "RIDValue")
                orgid_value.Value = TIMS.GetMyValue(Session(Cst_SearchName), "orgid_value")

                PlaceNo.Text = TIMS.GetMyValue(Session(Cst_SearchName), "PlaceNo")
                Place.Text = TIMS.GetMyValue(Session(Cst_SearchName), "Place")
                PageControler1.PageIndex = 0
                'PageControler1.PageIndex = TIMS.GetMyValue(Session(Cst_SearchName) , "PageIndex")
                MyValue = TIMS.GetMyValue(Session(Cst_SearchName), "PageIndex")
                If MyValue <> "" AndAlso IsNumeric(MyValue) Then
                    MyValue = CInt(MyValue)
                    PageControler1.PageIndex = MyValue
                End If
                MyValue = TIMS.GetMyValue(Session(Cst_SearchName), "submit")
                If MyValue = "1" Then
                    '有記憶值進入就查詢
                    Call sUtl_Search1()
                    'bt_search_Click(sender, e)
                End If
            End If
            Session(Cst_SearchName) = Nothing
            Return
        End If

        '第1次進入就查詢
        Call sUtl_Search1()
        'bt_search_Click(sender, e)
    End Sub

    '新增
    Private Sub bt_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_add.Click
        PlaceNo.Text = TIMS.ClearSQM(PlaceNo.Text)
        Place.Text = TIMS.ClearSQM(Place.Text)
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
        TIMS.SetMyValue(s_MyValue1, "PlaceNo", PlaceNo.Text)
        TIMS.SetMyValue(s_MyValue1, "Place", Place.Text)
        TIMS.SetMyValue(s_MyValue1, "ProcessType", "Insert")
        TIMS.SetMyValue(s_MyValue1, "RID", RIDValue.Value)
        TIMS.SetMyValue(s_MyValue1, "OrgID", orgid_value.Value)
        Dim url1 As String = "TC_01_013_add.aspx?ID=" & RqID1 & s_MyValue1
        Call TIMS.Utl_Redirect(Me, objconn, url1)
        'sm.RedirectUrlAfterBlock = ResolveUrl(url1)
    End Sub

    Sub KeepSearch()
        'Dim s_submit As String = If(DataGridTable.Visible, "1", "0")
        Session(Cst_SearchName) = ""
        Session(Cst_SearchName) &= "&prgid=" & TIMS.ClearSQM(Request("ID"))
        Session(Cst_SearchName) &= "&center=" & TIMS.ClearSQM(center.Text)
        Session(Cst_SearchName) &= "&RIDValue=" & TIMS.ClearSQM(RIDValue.Value)
        Session(Cst_SearchName) &= "&orgid_value=" & TIMS.ClearSQM(orgid_value.Value)
        Session(Cst_SearchName) &= "&PlaceNo=" & TIMS.ClearSQM(PlaceNo.Text)
        Session(Cst_SearchName) &= "&Place=" & TIMS.ClearSQM(Place.Text)
        Session(Cst_SearchName) &= "&PageIndex=" & DataGrid1.CurrentPageIndex + 1
        Session(Cst_SearchName) &= "&submit=" & If(DataGridTable.Visible, "1", "0")
    End Sub

    '查詢 search
    Sub sUtl_Search1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        PlaceNo.Text = TIMS.ClearSQM(PlaceNo.Text)
        Place.Text = TIMS.ClearSQM(Place.Text)
        txtAddress.Text = TIMS.ClearSQM(txtAddress.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        Dim sql As String = ""
        sql &= " SELECT P1.PTID" & vbCrLf 'PrimaryKey
        sql &= " ,P1.PlaceID ,P1.PlaceName" & vbCrLf
        sql &= " ,P1.Address" & vbCrLf
        sql &= " ,P1.FactMode" & vbCrLf
        sql &= " ,P1.PlacePic1 ,P1.PlacePic2" & vbCrLf
        sql &= " ,P1.ModifyType ,P1.ModifyDate " & vbCrLf
        sql &= " ,CONVERT(VARCHAR, p1.zipcode) + CONVERT(VARCHAR, iz.zipname) + CONVERT(VARCHAR, P1.Address) xzza1 " & vbCrLf
        sql &= " FROM ORG_ORGINFO O1 " & vbCrLf
        sql &= " JOIN PLAN_TRAINPLACE P1 ON O1.ComIDNO = P1.ComIDNO " & vbCrLf
        sql &= " LEFT JOIN VIEW_ZIPNAME iz ON iz.zipcode = P1.zipcode " & vbCrLf
        sql &= " WHERE 1=1 "
        '含已停用資料 ModifyType:D
        If Not chkboxDelData.Checked Then sql &= " AND ISNULL(P1.ModifyType,'') = '' " & vbCrLf

        If RIDValue.Value <> "" Then
            sql &= " AND O1.OrgID IN (SELECT OrgID FROM AUTH_RELSHIP WHERE RID = '" & RIDValue.Value & "') " & vbCrLf
        Else
            sql &= " AND O1.OrgID = '" & sm.UserInfo.OrgID & "' " & vbCrLf
        End If
        If PlaceNo.Text <> "" Then sql &= " AND P1.PLACEID LIKE '%" & PlaceNo.Text & "%' " & vbCrLf
        If Place.Text <> "" Then sql &= " AND P1.PLACENAME LIKE '%" & Place.Text & "%' " & vbCrLf
        If txtAddress.Text <> "" Then sql &= " AND CONVERT(VARCHAR, p1.zipcode) + CONVERT(VARCHAR, iz.zipname) + CONVERT(VARCHAR, P1.Address) LIKE '%" & txtAddress.Text & "%' " & vbCrLf
        sql &= " ORDER BY P1.ModifyDate DESC "
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料!!"
        DataGridTable.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Visible = True
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "PTID"
            'PageControler1.Sort = "PlaceID"
            PageControler1.Sort = "ModifyDate DESC, PlaceID ASC"  'edit，by:20181026
            PageControler1.ControlerLoad()
        End If
    End Sub

    '依機構設定訓練地點 Plan_TrainPlace
    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Call sUtl_Search1()
    End Sub

#Region "(No Use)"

    'Function Del_PlacePic(ByVal PTID As String)   '20090428之前的刪除功能
    '    Dim flag As Boolean = False
    '    Try
    '        Dim dt As DataTable
    '        Dim sql As String
    '        sql = "" & vbCrLf
    '        sql += "SELECT ptid, dbo.SUBSTR( dbo.SUBSTR(placepic1,-5,5),1,1) as depID, placepic1 as placepic1,placepic1 as okflag FROM Plan_TrainPlace WHERE placepic1 is not null and PTID='" & PTID & "'" & vbCrLf
    '        sql &= " union " & vbCrLf
    '        sql += "SELECT ptid, dbo.SUBSTR( dbo.SUBSTR(placepic2,-5,5),1,1) as depID, placepic2 as placepic1,placepic2 as okflag FROM Plan_TrainPlace WHERE placepic2 is not null and PTID='" & PTID & "'" & vbCrLf
    '        sql &= " order by placepic1 " & vbCrLf
    '        dt = DbAccess.GetDataTable(sql, objconn)
    '        If dt.Rows.Count > 0 Then
    '            For i As Int16 = 0 To dt.Rows.Count - 1
    '                Dim dr As DataRow = dt.Rows(i)
    '                If dr("placepic1").ToString <> "" Then
    '                    If File.Exists(Server.MapPath(Upload_Path & dr("placepic1").ToString)) Then
    '                        File.Delete(Server.MapPath(Upload_Path & dr("placepic1").ToString))
    '                    End If
    '                    'If strSyncUploadEnable = "1" Then
    '                    '    'Try
    '                    '    '    Dim ftp As New EnterpriseDT.Net.Ftp.FTPConnection
    '                    '    '    ftp.ServerAddress = strSyncUploadServer
    '                    '    '    ftp.ServerPort = strSyncUploadPort
    '                    '    '    ftp.UserName = strSyncUploadUser
    '                    '    '    ftp.Password = strSyncUploadPW
    '                    '    '    ftp.ServerDirectory = strSyncUploadFolder
    '                    '    '    ftp.Connect()
    '                    '    '    If ftp.Exists(Convert.ToString(dr("placepic1"))) = True Then
    '                    '    '        ftp.DeleteFile(Convert.ToString(dr("placepic1")))
    '                    '    '    End If
    '                    '    '    ftp.Close()
    '                    '    'Catch ex As Exception
    '                    '    '    'If ftp.IsConnected = True Then ftp.Close()
    '                    '    'End Try
    '                    'End If
    '                End If
    '            Next
    '        End If
    '        sql = "delete Plan_TrainPlace where PTID=" & PTID
    '        DbAccess.ExecuteNonQuery(sql, objconn)
    '        flag = True
    '    Catch ex As Exception
    '    End Try
    '    Return flag
    'End Function

#End Region

    '執行刪除功能/'還原
    Function RtnDel_PlacePic2(ByVal PTID As String, ByVal sRtnDel As String) As Boolean
        '刪除功能 20090428 改成 加一個欄位沒有真正的刪除
        Dim flag As Boolean = False '有異常狀況

        Dim u_SQL As String = ""
        Select Case sRtnDel
            Case cst_cmd_stop, cst_cmd_return
                Dim u_parms As New Hashtable
                u_parms.Clear()
                u_parms.Add("MODIFYTYPE", If(sRtnDel = cst_cmd_stop, "D", Convert.DBNull)) '啟用／停用
                u_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                u_parms.Add("PTID", Val(PTID))
                u_SQL = ""
                u_SQL &= " UPDATE PLAN_TRAINPLACE "
                u_SQL &= " SET MODIFYTYPE = @MODIFYTYPE ,MODIFYACCT = @MODIFYACCT ,MODIFYDATE = GETDATE() "
                u_SQL &= " WHERE 1=1 AND PTID = @PTID"
                DbAccess.ExecuteNonQuery(u_SQL, objconn, u_parms)

            Case cst_cmd_del
                TIMS.OpenDbConn(objconn)
                Dim oTrans As SqlTransaction = DbAccess.BeginTrans(objconn)
                Dim rst As Integer = 0
                Try
                    rst = DEL_TRAINPLACE(PTID, oTrans)
                Catch ex As Exception
                    TIMS.LOG.Warn(ex.Message, ex)
                    DbAccess.RollbackTrans(oTrans)
                    Dim strErrmsg As String = ""
                    strErrmsg &= "/*  DEL_TRAINPLACE */" & vbCrLf
                    strErrmsg &= "/*  ex.ToString: */" & vbCrLf
                    strErrmsg &= ex.ToString & vbCrLf
                    strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                    strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg)
                    Return flag
                End Try
                DbAccess.CommitTrans(oTrans)
        End Select
        flag = True '正常結束
        Return flag
    End Function

    Function DEL_TRAINPLACE(ByVal PTID As String, ByRef oTrans As SqlTransaction) As Integer
        Dim rst As Integer = 0
        'Dim sm As SessionModel = SessionModel.Instance()
        PTID = TIMS.ClearSQM(PTID)
        If PTID = "" Then Return 0 '異常不處理
        If Convert.ToString(sm.UserInfo.UserID) = "" Then Return 0 '異常不處理
        Dim parms As New Hashtable
        Dim sql As String = ""

        '確認
        sql = " SELECT * FROM PLAN_TRAINPLACE WHERE PTID = @PTID"
        parms.Clear()
        parms.Add("PTID", Val(PTID))
        Dim dt As DataTable = DbAccess.GetDataTable(sql, oTrans, parms)
        If dt.Rows.Count <> 1 Then Return 0 '異常不處理

        '變更人時
        sql = " UPDATE PLAN_TRAINPLACE SET MODIFYDATE = GETDATE() ,MODIFYACCT = @MODIFYACCT WHERE PTID= @PTID "
        parms.Clear()
        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        parms.Add("PTID", Val(PTID))
        DbAccess.ExecuteNonQuery(sql, oTrans, parms)

        Const cst_COLUMNS1 As String = " PTID,PLACEID,PLACENAME,COMIDNO,CONTACTNAME,CONTACTPHONE,CONTACTFAX,CONTACTEMAIL,CLASSIFICATION,FACTMODE,FACTMODEOTHER,ZIPCODE,ADDRESS,CONNUM,MASTERNAME,HWDESC,OTHERDESC,PLACEPIC1,PLACEPIC2,MODIFYACCT,MODIFYDATE,AREAPOSS,ZIP6W,MODIFYTYPE,ZIP_N,PINGNUMBER"
        '搬移資料
        sql = ""
        sql &= String.Format(" INSERT INTO PLAN_TRAINPLACEDELDATA({0}) ", cst_COLUMNS1)
        sql &= String.Format(" SELECT {0} FROM PLAN_TRAINPLACE ", cst_COLUMNS1)
        sql &= " WHERE PTID = @PTID AND MODIFYACCT = @MODIFYACCT "
        parms.Clear()
        parms.Add("PTID", Val(PTID))
        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        DbAccess.ExecuteNonQuery(sql, oTrans, parms)

        '刪除 
        sql = ""
        sql &= " DELETE PLAN_TRAINPLACE "
        sql &= " WHERE PTID = @PTID AND MODIFYACCT = @MODIFYACCT "
        parms.Clear()
        parms.Add("PTID", Val(PTID))
        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        DbAccess.ExecuteNonQuery(sql, oTrans, parms)
    End Function

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sPTID As String = TIMS.GetMyValue(e.CommandArgument, "PTID")
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        orgid_value.Value = TIMS.Get_OrgID(RIDValue.Value, objconn)
        If orgid_value.Value = "" OrElse orgid_value.Value = "-1" Then orgid_value.Value = sm.UserInfo.OrgID

        Dim RqID1 As String = TIMS.Get_MRqID(Me)
        Select Case e.CommandName
            Case cst_cmd_return '還原
                Dim rflag1 As Boolean = RtnDel_PlacePic2(sPTID, cst_cmd_return)
                If Not rflag1 Then
                    'Common.MessageBox(Me, "還原失敗!!")
                    lrMsg = "還原失敗!!"
                    sm.LastResultMessage = lrMsg
                    Exit Sub
                End If
                'Common.MessageBox(Me, "還原成功!!")
                lrMsg = "還原成功!!"
                sm.LastResultMessage = lrMsg
                Call sUtl_Search1() '還原成功 可查詢
            Case cst_cmd_edit
                KeepSearch()
                Dim s_MyValue1 As String = ""
                TIMS.SetMyValue(s_MyValue1, "PTID", sPTID)
                TIMS.SetMyValue(s_MyValue1, "ProcessType", "Update")
                TIMS.SetMyValue(s_MyValue1, "RID", RIDValue.Value)
                TIMS.SetMyValue(s_MyValue1, "OrgID", orgid_value.Value)
                Dim url1 As String = "TC_01_013_add.aspx?ID=" & RqID1 & s_MyValue1
                Call TIMS.Utl_Redirect(Me, objconn, url1)
            Case cst_cmd_stop '停用
                Dim rflag1 As Boolean = RtnDel_PlacePic2(sPTID, cst_cmd_stop)
                If Not rflag1 Then
                    'Common.MessageBox(Me, "刪除失敗!!")
                    lrMsg = "停用失敗!!"
                    sm.LastResultMessage = lrMsg
                    Exit Sub
                End If
                Call sUtl_Search1() '停用成功 可查詢
            Case cst_cmd_del '刪除
                If TIMS.ChkPTIDUse(sPTID, objconn) Then
                    lrMsg = "該場地目前被使用，不可刪除!!"
                    sm.LastResultMessage = lrMsg
                    Exit Sub
                End If
                Dim rflag1 As Boolean = RtnDel_PlacePic2(sPTID, cst_cmd_del)
                If Not rflag1 Then
                    'Common.MessageBox(Me, "刪除失敗!!")
                    lrMsg = "刪除失敗!!"
                    sm.LastResultMessage = lrMsg
                    Exit Sub
                End If
                'Common.MessageBox(Me, "刪除成功!!")
                lrMsg = "刪除成功!!"
                sm.LastResultMessage = lrMsg
                Call sUtl_Search1() '刪除成功 可查詢
            Case cst_cmd_view '檢視-for 訓練單位
                KeepSearch()
                Dim s_MyValue1 As String = ""
                TIMS.SetMyValue(s_MyValue1, "PTID", sPTID)
                TIMS.SetMyValue(s_MyValue1, "ProcessType", "View")
                TIMS.SetMyValue(s_MyValue1, "RID", RIDValue.Value)
                TIMS.SetMyValue(s_MyValue1, "OrgID", orgid_value.Value)

                Dim url1 As String = "TC_01_013_add.aspx?ID=" & RqID1 & s_MyValue1
                Call TIMS.Utl_Redirect(Me, objconn, url1)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim lbtView As LinkButton = e.Item.FindControl("lbtView")
                Dim lbtEdit As LinkButton = e.Item.FindControl("lbtEdit")
                Dim lbtReturn As LinkButton = e.Item.FindControl("lbtReturn")
                Dim lbtStop As LinkButton = e.Item.FindControl("lbtStop")
                Dim lbtDel As LinkButton = e.Item.FindControl("lbtDel")
                e.Item.Cells(0).Text = DataGrid1.CurrentPageIndex * DataGrid1.PageSize + e.Item.ItemIndex + 1
                lbtStop.Attributes("onclick") = "return confirm('確定要停用?');"
                lbtDel.Attributes("onclick") = "return confirm('確定要刪除?');"
                lbtReturn.Attributes("onclick") = "return confirm('確定要還原?');"
                'Dim but_edit, but_del As Button
                'but_edit = e.Item.Cells(7).FindControl("edit_but") '修改
                Dim cmdArg As String = ""
                TIMS.SetMyValue(cmdArg, "PTID", drv("PTID"))
                lbtStop.CommandArgument = cmdArg 'Convert.ToString(drv("PTID"))
                lbtDel.CommandArgument = cmdArg 'Convert.ToString(drv("PTID"))
                lbtView.CommandArgument = cmdArg 'Convert.ToString(drv("PTID"))
                lbtEdit.CommandArgument = cmdArg 'Convert.ToString(drv("PTID"))
                lbtReturn.CommandArgument = cmdArg '

                '分署／可刪、修、檢視、還原
                lbtDel.Style("display") = "" '永遠有刪

                lbtReturn.Style("display") = "none" '不能還原
                lbtView.Style("display") = "none"
                lbtStop.Style("display") = "" '可停用
                lbtEdit.Style("display") = ""
                If UCase(Convert.ToString(drv("ModifyType"))) = "D" Then
                    lbtReturn.Style("display") = "" '有還原
                    lbtView.Style("display") = "none"
                    lbtStop.Style("display") = "none" '不能停用
                    lbtEdit.Style("display") = ""
                End If

#Region "(No Use)"

                'Select Case sm.UserInfo.LID
                '    Case "2" '如果是委訓單位
                '        '2018-08-30 可修改/刪除-但有條件 
                '        lbtReturn.Style("display") = "none" '不能還原
                '        lbtView.Style("display") = "none" '不能檢視
                '        lbtStop.Style("display") = "" '提供刪除
                '        lbtEdit.Style("display") = "" '提供修改
                '        If UCase(Convert.ToString(drv("ModifyType"))) = "D" Then
                '            lbtReturn.Style("display") = "none" '不能還原
                '            lbtView.Style("display") = "" '僅檢視
                '            lbtStop.Style("display") = "none" '不能停用
                '            lbtEdit.Style("display") = "none" '不能修改
                '        End If
                'End Select

#End Region

                '有停用 就沒有還原
                lbtStop.Visible = True
                lbtReturn.Visible = False
                If UCase(Convert.ToString(drv("ModifyType"))) = "D" Then
                    '有還原 就沒停用 
                    lbtStop.Visible = False
                    lbtReturn.Visible = True
                End If
                Dim sFactMode As String = "其他"
                Select Case Convert.ToString(drv("FactMode"))
                    Case "1"
                        sFactMode = "教室"
                    Case "2"
                        sFactMode = "演講廳"
                    Case "3"
                        sFactMode = "會議室"
                End Select
                e.Item.Cells(4).Text = sFactMode
        End Select
    End Sub
End Class