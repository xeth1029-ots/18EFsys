Partial Class Chg_OrgName
    Inherits AuthBasePage

    'Dim Old_OrgN As String
    'Dim trans As SqlTransaction
    'Dim CommandArgument As String

#Region "(No Use)"

    'Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
    '    '20080813 andy 還原及查詢功能不開放給使用者
    '    'Dim restore As Button = e.Item.FindControl("restore")
    '    'Dim drv As DataRowView = e.Item.DataItem
    '    'Dim dt As DataTable
    '    'Dim dr, dr2 As DataRow
    '    'Dim sql, id As String
    '    'Dim ArrayEditID As Integer
    '    'Dim FirstEditID As New ArrayList
    '    'sql = "select   distinct   years      from   Org_OrgNameHistory"
    '    'dt = DbAccess.GetDataTable(sql)

    '    'For Each dr In dt.Rows   '取每年度的最後一筆變更記錄
    '    '    sql = "select top  1  EditID   from  Org_OrgNameHistory  "
    '    '    sql &= " where  years = '" & dr("years") & "'"
    '    '    sql &= " order by   ModifyDate  desc  "
    '    '    FirstEditID.Add(DbAccess.ExecuteScalar(sql))
    '    'Next

    '    'If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
    '    '    restore.Enabled = False
    '    '    For Each ArrayEditID In FirstEditID
    '    '        If ArrayEditID = drv("EditID") Then
    '    '            restore.Enabled = True
    '    '        End If
    '    '    Next
    '    '    restore.CommandArgument = drv("EditID").ToString
    '    'End If
    'End Sub

    'Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
    '    '20080813 andy 還原及查詢功能不開放給使用者
    '    'Select Case e.CommandName
    '    '    Case "restore"
    '    '        Dim sql, strScript As String
    '    '        conn = DbAccess.GetConnection
    '    '        conn.Open()
    '    '        trans = DbAccess.BeginTrans(conn)
    '    '        Try
    '    '            sql &= " update  Org_OrgInfo "
    '    '            sql &= " set  OrgName= (SELECT  OrgName  from   Org_OrgNameHistory  where EditID=" & e.CommandArgument & ")"
    '    '            sql &= " where  orgid = '" & Request("orgid") & "'"
    '    '            DbAccess.ExecuteNonQuery(sql, trans)
    '    '            DbAccess.CommitTrans(trans)
    '    '        Catch ex As Exception
    '    '            DbAccess.RollbackTrans(trans)
    '    '            conn.Close()
    '    '            Throw ex
    '    '        End Try
    '    '        strScript = "<script>alert('機構名稱還原完成!!!');"
    '    '        strScript += "</script>"
    '    '        Page.RegisterStartupScript("", strScript)
    '    '        conn.Close()
    '    '        GetOrgName()
    '    '        showOrgHis()
    '    'End Select
    'End Sub

    'Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
    '    '20080813 andy 還原及查詢功能不開放給使用者
    '    'Dim strScript As String
    '    'If Not lb_years.SelectedItem Is Nothing Then
    '    '    TextBox1.Text = "You Selected" & " " & lb_years.SelectedItem.Value
    '    '    showOrgHis(lb_years.SelectedItem.Value.ToString())
    '    'Else
    '    '    'strScript = "<script>alert('尚未選取欲查詢年度!!!');"
    '    '    'strScript += "</script>"
    '    '    'Page.RegisterStartupScript("", strScript)
    '    '    showOrgHis()
    '    'End If
    'End Sub

#End Region

    Dim rqOrgID As String = ""
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        Hid_orgid.Value = TIMS.sUtl_GetRqValue(Me, "orgid", Hid_orgid.Value)
        If Not IsPostBack Then
            Me.bt_save.Attributes.Add("Onclick", "javascript:checkData( ); return confirm('是否確定要儲存?');")
            Call Get_ORGNAME_OLD() '取得機構名稱-變更歷史
            Call showOrgHis("")
        End If
    End Sub

    ''' <summary> 取得機構名稱-變更歷史 </summary>
    Sub Get_ORGNAME_OLD()
        Dim UserID As String = sm.UserInfo.UserID
        Hid_orgid.Value = TIMS.ClearSQM(Hid_orgid.Value)
        If Hid_orgid.Value = "" Then Exit Sub
        Dim orgid As String = Hid_orgid.Value

        Dim parms As New Hashtable
        parms.Add("orgid", orgid)

        Dim sql As String = " SELECT MAX(Years) Years FROM Org_OrgNameHistory WHERE orgid =@orgid "

        Dim MaxYear As String = Convert.ToString(DbAccess.ExecuteScalar(sql, objconn, parms))

        sql = "SELECT OrgName FROM Org_OrgNameHistory WHERE orgid =@orgid "
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        Dim v_Old_OrgName As String = ""
        '檢查是否於Org_OrgNameHistory內有變更資料；(1) 若無資料or (2) 有1筆資料or (3) 選取變更年度=變更歷史記錄最後一年,則取Org_OrgInfo內的OrgName當作原始單位名稱
        If dt.Rows.Count = 0 OrElse dt.Rows.Count = 1 OrElse (Convert.ToString(Me.sm.UserInfo.Years) = MaxYear) Then
            sql = " SELECT OrgName FROM Org_OrgInfo WHERE orgid = @orgid "
            v_Old_OrgName = DbAccess.ExecuteScalar(sql, objconn, parms)
        Else
            '有多筆且年度不相同。
            sql = " SELECT OrgName FROM Org_OrgNameHistory WHERE 1=1 AND orgid = @orgid AND  OrgName != '' ORDER BY YEARS DESC ,MODIFYDATE DESC "
            dt = DbAccess.GetDataTable(sql, objconn, parms)

            '若選取年度查詢有歷史變更記錄則取該年度最後一筆記錄為變更單位名稱
            v_Old_OrgName = If(dt.Rows.Count > 0, dt.Rows(0)("OrgName"), "(意外查無資料。)")
        End If
        Old_OrgName.Text = v_Old_OrgName
    End Sub

    '儲存
    Private Sub bt_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_save.Click
        New_OrgName.Text = TIMS.ClearSQM(New_OrgName.Text)
        Old_OrgName.Text = TIMS.ClearSQM(Old_OrgName.Text)
        If New_OrgName.Text = "" Then
            msg_1.Text = "異動後機構名稱 不可為空!"
            Return ' Exit Sub
        End If
        If Old_OrgName.Text = "" Then
            msg_1.Text = "目前機構名稱 不可為空!"
            Return ' Exit Sub
        End If
        If New_OrgName.Text = Old_OrgName.Text Then
            msg_1.Text = "異動後機構名稱 不可與 目前機構名稱 相同!"
            Return ' Exit Sub
        End If
        Dim oErrmsg As String = ""
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "Years", Convert.ToString(sm.UserInfo.Years))
        TIMS.SetMyValue2(htSS, "NewOrgName", New_OrgName.Text)
        TIMS.SetMyValue2(htSS, "OldOrgName", Old_OrgName.Text)
        TIMS.SetMyValue2(htSS, "OrgID", Hid_orgid.Value)
        Dim rst As Boolean = ADD_ORGNAMEHIS2(htSS, objconn, oErrmsg)
        If oErrmsg <> "" Then
            msg_1.Text = String.Concat("機構異動有誤!", oErrmsg)
            Return 'Exit Sub
        End If
        If Not rst Then
            'Common.MessageBox(Me, "機構異動有誤!!")
            msg_1.Text = "機構異動有誤!!"
            Return 'Exit Sub
        End If

        Call Get_ORGNAME_OLD()
        Call showOrgHis("")
        'strScript = "<script>alert('機構名稱異動完成!!!');window.close(); window.opener.location.reload();  "
        'strScript += "opener.form1.TBtitle.value='" & New_OrgName.Text & "';window.close(); </script> "
        'strScript += "opener.location.href='TC_01_002.aspx?ID=" & Request("ID") & "';window.close(); </script>"
        'strScript += "opener.location.href='" & Request("BackPage") & "?ID=" & Request("ID") & "';window.close(); </script>"
        'strScript += " </script>"
        'strScript += "opener.location.href='TC_01_002_add.aspx?ID=ProcessType=Share&orgid=" & Request("orgid") & "&distid=" & Request("distid") & "&planid=" & Request("planid") & "&rid=" & Request("rid") & "&ID= " & Request("ID") & " ';window.close(); </script>"
        Dim RqID1 As String = TIMS.Get_MRqID(Me)
        Dim rq_BackPage As String = TIMS.sUtl_GetRqValue(Me, "BackPage")
        Dim strScript As String = ""
        strScript = "<script>alert('機構名稱異動完成!!!'); opener.location.href='" & rq_BackPage & "?ID=" & RqID1 & "';</script>"
        Page.RegisterStartupScript(TIMS.xBlockName, strScript)
    End Sub


    Sub showOrgHis(ByVal Search_Year As String)
        Hid_orgid.Value = TIMS.ClearSQM(Hid_orgid.Value)
        If Hid_orgid.Value = "" Then Exit Sub

        Dim parms As New Hashtable From {{"ORGID", Val(Hid_orgid.Value)}}
        If Search_Year <> "" Then parms.Add("YEARS", Search_Year)

        Dim sql As String = ""
        sql &= " SELECT a.EditID,a.Years ,a.ORGNAME ,b.Name ModifyName " & vbCrLf
        sql &= " ,format(a.ModifyDate,'yyyy-MM-dd HH:mm:ss') ModifyDate" & vbCrLf
        sql &= " FROM ORG_ORGNAMEHISTORY a " & vbCrLf
        sql &= " LEFT JOIN AUTH_ACCOUNT b ON a.ModifyAcct = b.account " & vbCrLf
        sql &= " WHERE a.ORGID =@ORGID " & vbCrLf
        If Search_Year <> "" Then sql &= " AND a.YEARS =@YEARS"
        sql &= " ORDER BY a.EditID desc ,a.Years " & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        DataGrid1.Visible = False
        msg_1.Text = "查無資料!!"
        If dt.Rows.Count = 0 Then Exit Sub

        DataGrid1.Visible = True
        msg_1.Text = ""
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()

        If msg_1.Text = "" Then msg_1.Text = "若有變更，請重新進入此功能，才能看到變更結果!!"

        'sql = "  select   distinct  Years   from  Org_OrgNameHistory  order by    Years desc "
        'dt = DbAccess.GetDataTable(sql)
        'lb_years.Items.Clear()
        'For Each dr In dt.Rows   '取每年度的最後一筆變更記錄
        '    lb_years.Items.Add(dr("Years"))
        'Next
    End Sub

    Function ADD_ORGNAMEHIS2(ByRef htSS As Hashtable, ByVal objconn As SqlConnection, ByRef oErrmsg As String) As Boolean
        Dim rst As Boolean = False

        Dim sMaxYear As String = TIMS.GetMyValue2(htSS, "Years")
        Dim NewOrgName As String = TIMS.GetMyValue2(htSS, "NewOrgName")
        Dim OldOrgName As String = TIMS.GetMyValue2(htSS, "OldOrgName")
        Dim OrgID As String = TIMS.GetMyValue2(htSS, "OrgID")
        'NewOrgName = TIMS.ClearSQM(NewOrgName) 'sm.UserInfo.Years
        If Convert.ToString(sm.UserInfo.Years) = "" Then oErrmsg = "(登入年度有誤)！" : Return False
        If Convert.ToString(sm.UserInfo.UserID) = "" Then oErrmsg = "(登入資訊有誤)！" : Return False
        If NewOrgName = "" OrElse OldOrgName = "" OrElse OrgID = "" Then oErrmsg = "(輸入資料有誤)！(BEFE)" : Return False

        Call TIMS.OpenDbConn(objconn)
        'select years ,count(1) cnt from Org_OrgNameHistory group by years order by 1
        'New_OrgName.Text = TIMS.ClearSQM(New_OrgName.Text)

        '10分鐘內不可再執行該應用
        Dim sql3 As String = " SELECT 1 FROM ORG_ORGNAMEHISTORY WHERE ORGID = @ORGID AND MODIFYDATE >= GETDATE()-(1.0/24/6)"
        Using sCmd3 As New SqlCommand(sql3, objconn)
            Dim dt3 As New DataTable
            With sCmd3
                .Parameters.Clear()
                .Parameters.Add("ORGID", SqlDbType.Int).Value = Val(OrgID)
                dt3.Load(.ExecuteReader())
            End With
            If dt3.Rows.Count > 0 Then
                oErrmsg = "機構名稱變動頻繁,停止變動!(10分鐘內已有變動資料)"
                Return False
            End If
        End Using

        Dim dt As New DataTable
        Dim sql As String = " SELECT 'x' FROM ORG_ORGNAMEHISTORY WHERE ORGID=@ORGID"
        Using sCmd As New SqlCommand(sql, objconn)
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("ORGID", SqlDbType.Int).Value = Val(OrgID)
                dt.Load(.ExecuteReader())
            End With
        End Using

        Dim i_sql As String = ""
        i_sql &= " INSERT INTO ORG_ORGNAMEHISTORY (EDITID ,Years ,ORGID ,OrgName ,ModifyAcct ,ModifyDate) "
        i_sql &= " VALUES (@EDITID ,@Years ,@ORGID ,@OrgName ,@ModifyAcct ,GETDATE()) "
        Dim u_sql As String = ""
        u_sql &= " UPDATE Org_OrgInfo SET OrgName = @OrgName ,ModifyAcct = @ModifyAcct ,ModifyDate = GETDATE() WHERE ORGID = @ORGID "

        Dim iEDITID As Integer = 0
        If dt.Rows.Count = 0 Then
            '當變更歷史資料無記錄時 1.新增1筆變更記錄 2.直接更新目前Org_OrgInfo內的OrgName
            iEDITID = DbAccess.GetNewId(objconn, "ORG_ORGNAMEHISTORY_EDITID_SEQ,ORG_ORGNAMEHISTORY,EDITID")

            Dim i_parms As New Hashtable From {{"EDITID", iEDITID}, {"Years", sm.UserInfo.Years}, {"ORGID", Val(OrgID)},
                {"OrgName", OldOrgName}, {"ModifyAcct", sm.UserInfo.UserID}}
            DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)

            Dim u_parms As New Hashtable From {{"OrgName", NewOrgName}, {"ModifyAcct", sm.UserInfo.UserID}, {"ORGID", Val(OrgID)}}
            DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)

            Return True
        Else
            Dim sql2 As String = " SELECT MAX(Years) MAXYEARS FROM ORG_ORGNAMEHISTORY WHERE ORGID = @ORGID "
            Using sCmd2 As New SqlCommand(sql2, objconn)
                '如果是當年度第1筆資料同時向前查詢亦無異動記錄時，
                '則將 目前欲新增的單位名稱 <==> 最原始的單位名稱 (名稱、modifydate) 兩者互換
                With sCmd2
                    .Parameters.Clear()
                    .Parameters.Add("ORGID", SqlDbType.Int).Value = Val(OrgID)
                    sMaxYear = Convert.ToString(.ExecuteScalar())
                End With
            End Using

            If sMaxYear <> "" AndAlso Val(sm.UserInfo.Years) >= Val(sMaxYear) Then
                iEDITID = DbAccess.GetNewId(objconn, "ORG_ORGNAMEHISTORY_EDITID_SEQ,ORG_ORGNAMEHISTORY,EDITID")
                Dim i_parms As New Hashtable From {
                    {"EDITID", iEDITID},
                    {"Years", sm.UserInfo.Years},
                    {"ORGID", Val(OrgID)},
                    {"OrgName", OldOrgName},
                    {"ModifyAcct", sm.UserInfo.UserID}
                }
                DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)
            Else
                '改舊年度的資料 (只需新增1筆不須要使用新機構名稱)
                iEDITID = DbAccess.GetNewId(objconn, "ORG_ORGNAMEHISTORY_EDITID_SEQ,ORG_ORGNAMEHISTORY,EDITID")
                Dim i_parms As New Hashtable From {
                    {"EDITID", iEDITID},
                    {"Years", sm.UserInfo.Years},
                    {"ORGID", Val(OrgID)},
                    {"OrgName", NewOrgName}, 'OldOrgName
                    {"ModifyAcct", sm.UserInfo.UserID}
                }
                DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)
            End If

            If sMaxYear <> "" AndAlso Val(sm.UserInfo.Years) >= Val(sMaxYear) Then
                Dim u_parms As New Hashtable From {
                    {"OrgName", NewOrgName},
                    {"ModifyAcct", sm.UserInfo.UserID},
                    {"ORGID", Val(OrgID)}
                }
                DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)
                Return True
            End If
        End If

        Return rst
    End Function

#Region "(No Use)"

    'Sub add_OrgNameHis()
    '    New_OrgName.Text = TIMS.ClearSQM(New_OrgName.Text)
    '    Call TIMS.OpenDbConn(objconn)
    '    Dim EDITID As Integer = 0
    '    Try
    '        Dim sql As String = "select max( Years)  maxYear  from Org_OrgNameHistory  where orgid = " & orgid & ""
    '        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
    '        If IsDBNull(dr("maxYear")) Then  '當變更歷史資料無記錄時 1.新增一筆變更記錄 2.直接更新目前Org_OrgInfo內的OrgName
    '            EDITID = DbAccess.GetNewId(objconn, "ORG_ORGNAMEHISTORY_EDITID_SEQ,ORG_ORGNAMEHISTORY,EDITID")
    '            sql = ""
    '            sql &= " insert  into Org_OrgNameHistory (EDITID,Years,OrgID,OrgName,ModifyAcct,ModifyDate)"
    '            sql &= "  values(" & EDITID & ",'" & sm.UserInfo.Years & "','" & orgid & "','" & Old_OrgName.Text & "','" & sm.UserInfo.UserID & "' ,getdate()) "
    '            DbAccess.ExecuteNonQuery(sql, objconn)
    '            sql = ""
    '            sql &= " update Org_OrgInfo "
    '            sql &= " set  OrgName='" & New_OrgName.Text & "',"
    '            sql &= " ModifyAcct='" & sm.UserInfo.UserID & "',"
    '            sql &= " ModifyDate=getdate()"
    '            sql &= " where orgid = '" & Request("orgid") & "'"
    '            DbAccess.ExecuteNonQuery(sql, objconn)
    '        Else
    '            If Not IsDBNull(dr("maxYear")) Then
    '                If Convert.ToInt16(sm.UserInfo.Years) >= Convert.ToInt16(dr("maxYear")) Then
    '                    EDITID = DbAccess.GetNewId(objconn, "ORG_ORGNAMEHISTORY_EDITID_SEQ,ORG_ORGNAMEHISTORY,EDITID")
    '                    sql = ""
    '                    sql &= " insert into Org_OrgNameHistory (EDITID,Years,OrgID,OrgName,ModifyAcct,ModifyDate)"
    '                    sql &= " values(" & EDITID & ",'" & sm.UserInfo.Years & "','" & orgid & "','" & Old_OrgName.Text & "','" & sm.UserInfo.UserID & "' ,getdate()) "
    '                    DbAccess.ExecuteNonQuery(sql, objconn)
    '                    sql = ""
    '                    sql &= " update  Org_OrgInfo "
    '                    sql &= " set  OrgName='" & New_OrgName.Text & "',"
    '                    sql &= " ModifyAcct='" & sm.UserInfo.UserID & "',"
    '                    sql &= " ModifyDate=getdate()"
    '                    sql &= " where  orgid = '" & Request("orgid") & "'"
    '                    DbAccess.ExecuteNonQuery(sql, objconn)
    '                Else
    '                    '如果是當年度第一筆資料同時向前查詢亦無異動記錄時，
    '                    '則將 目前欲新增的單位名稱 <==> 最原始的單位名稱 (名稱、modifydate) 兩者互換
    '                    If Me.Session("newRecordInYear") = "Y" Then
    '                        sql = "  select ROWID,A.* from Org_OrgNameHistory A"
    '                        sql &= " where  0=0  and  A.orgid=" & orgid & " order  by  A.ModifyDate  "
    '                        dr = DbAccess.GetOneRow(sql, objconn)
    '                        ' sql &= " set  Years='" & sm.UserInfo.Years & "',"
    '                        sql = ""
    '                        sql &= " update  Org_OrgNameHistory  "   '更新異動記錄內最原始單位名稱
    '                        sql &= " set OrgID='" & orgid & "'"
    '                        sql &= " ,OrgName='" & New_OrgName.Text & "'"
    '                        sql &= " ,ModifyAcct='" & sm.UserInfo.UserID & "' "
    '                        sql &= " where ROWID= '" & Convert.ToString(dr("ROWID")) & "'"
    '                        sql &= " and EditID='" & dr("EditID") & "' "
    '                        DbAccess.ExecuteNonQuery(sql, objconn)
    '                        EDITID = DbAccess.GetNewId(objconn, "ORG_ORGNAMEHISTORY_EDITID_SEQ,ORG_ORGNAMEHISTORY,EDITID")
    '                        sql = ""
    '                        sql &= " insert  into  Org_OrgNameHistory (EDITID,Years,OrgID,OrgName,ModifyAcct,ModifyDate) "
    '                        sql += "values( " & EDITID & ",'" & sm.UserInfo.Years & "','" & dr("OrgID") & "','" & dr("OrgName") & "','" & sm.UserInfo.UserID & "' ,getdate()) "
    '                        DbAccess.ExecuteNonQuery(sql, objconn)
    '                    Else
    '                        EDITID = DbAccess.GetNewId(objconn, "ORG_ORGNAMEHISTORY_EDITID_SEQ,ORG_ORGNAMEHISTORY,EDITID")
    '                        sql &= " insert  into Org_OrgNameHistory (EDITID,Years,OrgID,OrgName,ModifyAcct,ModifyDate)"
    '                        sql &= "  values(" & EDITID & ", '" & sm.UserInfo.Years & "','" & orgid & "','" & New_OrgName.Text & "','" & sm.UserInfo.UserID & "' ,getdate()) "
    '                        DbAccess.ExecuteNonQuery(sql, objconn)
    '                    End If
    '                End If
    '            End If
    '        End If

    '        'DbAccess.ExecuteNonQuery(sql, trans)
    '        'DbAccess.CommitTrans(trans)
    '        'Call TIMS.CloseDbConn(conn)
    '        Me.Session("newRecordInYear") = "N"
    '        Call GetOrgName()
    '        Call showOrgHis()
    '        Dim strScript As String = ""
    '        'strScript = "<script>alert('機構名稱異動完成!!!');window.close(); window.opener.location.reload();  "
    '        strScript = "<script>alert('機構名稱異動完成!!!'); opener.location.href='" & Request("BackPage") & "?ID=" & Request("ID") & "'; </script>"
    '        'strScript += "opener.form1.TBtitle.value='" & New_OrgName.Text & "';window.close(); </script> "
    '        'strScript += "opener.location.href='TC_01_002.aspx?ID=" & Request("ID") & "';window.close(); </script>"
    '        'strScript += "opener.location.href='" & Request("BackPage") & "?ID=" & Request("ID") & "';window.close(); </script>"
    '        'strScript += " </script>"
    '        'strScript += "opener.location.href='TC_01_002_add.aspx?ID=ProcessType=Share&orgid=" & Request("orgid") & "&distid=" & Request("distid") & "&planid=" & Request("planid") & "&rid=" & Request("rid") & "&ID= " & Request("ID") & " ';window.close(); </script>"
    '        Page.RegisterStartupScript("", strScript)
    '    Catch ex As Exception
    '        'DbAccess.RollbackTrans(trans)
    '        Throw ex
    '    End Try
    'End Sub

#End Region

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        DbAccess.CloseDbConn(objconn)
        Common.RespWrite(Me, "<Script>window.close();</Script>")
    End Sub
End Class