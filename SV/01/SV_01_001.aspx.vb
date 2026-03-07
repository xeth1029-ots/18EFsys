Partial Class SV_01_001
    Inherits AuthBasePage

    'Dim FunDr As DataRow
    Const cst_Insert As String = "I" 'uType Insert/Update
    Const cst_Update As String = "E" 'uType Insert/Update

    Dim dtKSK As DataTable = Nothing
    Dim dtSdS As DataTable = Nothing
    Dim ff As String = ""
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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        '分頁設定 Start
        'PageControler1 = Me.FindControl("PageControler1")
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If sm.UserInfo.FunDt Is Nothing Then
            Common.RespWrite(Me, "<script>alert('Session過期');</script>")
            Common.RespWrite(Me, "<script>Top.location.href='../../logout.aspx';</script>")
        Else
            Dim FunDt As DataTable = sm.UserInfo.FunDt
            Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
            If FunDrArray.Length = 0 Then
                Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
                Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
            Else
                'Dim FunDr As DataRow = FunDrArray(0)
            End If
        End If

        If Not IsPostBack Then
            ddlSType = TIMS.Get_SurveyType(ddlSType, objconn)
            'Common.SetListItem(ddlSType, "01") '預設

            table_F.Visible = True
            table_I.Visible = False
            PageControler1.Visible = False
        End If

        'Me.TRddlSType.Style("display") = "inline"
        'Me.TRddlSurveyType.Style("display") = "inline"
        Me.TRddlSType.Style("display") = "none"
        Me.TRddlSurveyType.Style("display") = "none"

    End Sub

    '查詢
    Sub dt_search()
        'Dim dt As DataTable
        'Dim str As String
        Ipt_Name.Value = TIMS.ClearSQM(Ipt_Name.Value)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select a.SVID" & vbCrLf
        sql += " ,a.Name" & vbCrLf
        sql += " ,dbo.DECODE(a.Avail, 'Y','啟用' ,'不啟用') Avail" & vbCrLf
        sql += " ,a.Avail ISUSE" & vbCrLf
        sql += " ,a.internal " & vbCrLf
        sql += " ,st.STID" & vbCrLf
        sql += " ,st.STName" & vbCrLf
        sql += " from ID_Survey a" & vbCrLf
        sql += " LEFT JOIN ID_SurveyTypeRel sr on sr.SVID=a.SVID" & vbCrLf
        sql += " LEFT JOIN Key_SurveyType st on st.STID=sr.STID" & vbCrLf
        sql += " where 1=1" & vbCrLf
        If Ipt_Name.Value <> "" Then   '搜尋條件
            sql += " and a.Name like '%" & Ipt_Name.Value & "%' " & vbCrLf
        End If
        If ddlSType.SelectedValue <> "" Then   '搜尋條件
            sql += " and st.STID= '" & ddlSType.SelectedValue & "' " & vbCrLf
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料"
        msg.Visible = True
        table_F.Visible = True
        table_I.Visible = False
        DataGrid1.Visible = False
        PageControler1.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            'msg.Visible = False
            table_F.Visible = True
            table_I.Visible = False
            DataGrid1.Visible = True
            PageControler1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    '新增
    Function UpdateInsertSurvey(ByVal uType As String, ByVal SS As String) As Integer
        Dim rst As Integer = 0

        Dim SurveyName As String = TIMS.GetMyValue(SS, "SurveyName")
        Dim ISUSE2 As String = TIMS.GetMyValue(SS, "ISUSE2")
        Dim CHKinternalVAL As String = TIMS.GetMyValue(SS, "CHKinternalVAL")
        Dim ValSVID As String = TIMS.GetMyValue(SS, "ValSVID")

        Dim sql As String = ""
        sql = ""
        sql &= " INSERT INTO ID_Survey (SVID,Name,Avail,internal,ModifyAcct,modifydate)"
        sql += " values (@SVID,@Name,@Avail,@internal,@ModifyAcct,getdate())"
        Dim iCmd As New SqlCommand(sql, objconn)

        sql = ""
        sql &= " UPDATE ID_Survey"
        sql += " SET Name =@Name"
        sql += " ,Avail =@Avail"
        sql += " ,internal =@internal"
        sql += " ,ModifyAcct =@ModifyAcct"
        sql += " ,ModifyDate =getdate()"
        sql += " Where SVID = @SVID "
        Dim uCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        Select Case uType
            Case cst_Insert
                Dim iSVID As Integer = DbAccess.GetNewId(objconn, "ID_SURVEY_SVID_SEQ,ID_SURVEY,SVID")
                With iCmd
                    .Parameters.Clear()
                    .Parameters.Add("SVID", SqlDbType.Int).Value = iSVID
                    .Parameters.Add("Name", SqlDbType.VarChar).Value = SurveyName
                    .Parameters.Add("Avail", SqlDbType.VarChar).Value = IIf(ISUSE2 = "Y", "Y", "N")
                    .Parameters.Add("internal", SqlDbType.VarChar).Value = IIf(CHKinternalVAL = "Y", "Y", Convert.DBNull)
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .ExecuteNonQuery()
                End With
                rst = iSVID
            Case cst_Update
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("Name", SqlDbType.VarChar).Value = SurveyName 'IputQName.Value
                    .Parameters.Add("Avail", SqlDbType.VarChar).Value = IIf(ISUSE2 = "Y", "Y", "N")
                    .Parameters.Add("internal", SqlDbType.VarChar).Value = IIf(CHKinternalVAL = "Y", "Y", Convert.DBNull)
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .Parameters.Add("SVID", SqlDbType.VarChar).Value = ValSVID 'HidSVID.Value
                    .ExecuteNonQuery()
                End With
                rst = Val(ValSVID) 'Val(HidSVID.Value)
        End Select

        Return rst
    End Function

    '顯示修改資料
    Sub loaddata1(ByVal ssSVID As String)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select	a.SVID" & vbCrLf
        sql += " ,a.Name" & vbCrLf
        sql += " ,dbo.DECODE(a.Avail, 'Y','啟用' ,'不啟用') Avail" & vbCrLf
        sql += " ,a.Avail ISUSE" & vbCrLf
        sql += " ,a.internal" & vbCrLf
        sql += " ,st.STID" & vbCrLf
        sql += " ,st.STName" & vbCrLf
        sql += " from ID_Survey a" & vbCrLf
        sql += " LEFT JOIN ID_SurveyTypeRel sr on sr.SVID=a.SVID" & vbCrLf
        sql += " LEFT JOIN Key_SurveyType st on st.STID=sr.STID" & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " AND a.SVID = @SVID " & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SVID", SqlDbType.VarChar).Value = ssSVID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            IputQName.Value = dr("Name")
            Common.SetListItem(ddlSurveyType, dr("STID"))

            ISUSE.Checked = False
            If Convert.ToString(dr("ISUSE")) = "Y" Then
                ISUSE.Checked = True
            End If

            IputQName.Disabled = False '.Enabled = True
            chkinternal.Enabled = True
            chkinternal.Checked = False
            If Convert.ToString(dr("internal")) = "Y" Then
                IputQName.Disabled = True 'Enabled = False
                chkinternal.Enabled = False
                chkinternal.Checked = True
            End If
        End If
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand

        Dim sCmdArg As String = e.CommandArgument
        HidSVID.Value = TIMS.GetMyValue(sCmdArg, "SVID")
        If HidSVID.Value = "" Then Exit Sub

        Select Case e.CommandName
            Case "edit"        '修改
                ddlSurveyType = TIMS.Get_SurveyType(ddlSurveyType, objconn)
                Common.SetListItem(ddlSurveyType, "01") '預設
                HidMode.Value = "E"

                table_I.Visible = True
                table_F.Visible = False
                DataGrid1.Visible = False
                PageControler1.Visible = False

                Call loaddata1(HidSVID.Value)

            Case "del"          '刪除動作開始
                dtKSK = TIMS.Get_dtKSK(HidSVID.Value, objconn)
                If dtKSK.Select(ff).Length <> 0 Then
                    Common.MessageBox(Me, "尚有資料問卷分類，無法刪除")
                    Exit Sub
                End If
                dtSdS = TIMS.Get_dtSdS(HidSVID.Value, objconn)
                If dtSdS.Select(ff).Length <> 0 Then
                    Common.MessageBox(Me, "【問卷資料填寫】已有學員填寫的資料，無法刪除")
                    Exit Sub
                End If

                Dim sql As String = ""
                sql = "DELETE ID_Survey WHERE SVID='" & HidSVID.Value & "'"
                DbAccess.ExecuteNonQuery(sql, objconn)
                sql = "DELETE ID_SurveyTypeRel WHERE SVID='" & HidSVID.Value & "'"
                DbAccess.ExecuteNonQuery(sql, objconn)
                Common.MessageBox(Me, "刪除成功")

                dt_search()
        End Select
    End Sub



    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.EditItem, ListItemType.AlternatingItem
                Dim Btn_edit As Button = e.Item.FindControl("Btn_edit")
                Dim Btn_del As Button = e.Item.FindControl("Btn_del")
                Dim drv As DataRowView = e.Item.DataItem

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "SVID", Convert.ToString(drv("SVID")))
                Btn_edit.CommandArgument = sCmdArg 'drv("SVID").ToString
                Btn_del.CommandArgument = sCmdArg 'drv("SVID").ToString
                Btn_del.Attributes("onclick") = TIMS.cst_confirm_delmsg1

                dtKSK = TIMS.Get_dtKSK(Convert.ToString(drv("SVID")), objconn)
                If Btn_del.Enabled AndAlso dtKSK.Select(ff).Length <> 0 Then
                    Btn_del.CommandArgument = ""
                    Btn_del.Enabled = False
                    Btn_del.ToolTip = "此【問卷名稱】底下的【問卷分類標題設定】有資料，不能刪除，若執意要刪除，請先刪除【問卷分類標題設定】!!"
                End If

                dtSdS = TIMS.Get_dtSdS(Convert.ToString(drv("SVID")), objconn)
                If Btn_del.Enabled AndAlso dtSdS.Select(ff).Length <> 0 Then
                    Btn_edit.Enabled = False
                    Btn_edit.ToolTip = "此【問卷名稱】底下的【問卷資料填寫】已有學員填寫的資料，不能修改，若執意要修改，請先刪除【問卷資料填寫】!!"
                End If
                If Convert.ToString(drv("internal")) = "Y" Then
                    Btn_del.CommandArgument = ""
                    Btn_del.Enabled = False
                    Btn_del.ToolTip = "內部使用，不能刪除!!"
                End If
        End Select


    End Sub


#Region "NO USE"
    'Function Get_dtKSK(ByVal SVID As String, ByRef oConn As SqlConnection) As DataTable
    '    Dim rst As New DataTable
    '    Dim sql As String = ""
    '    sql = "SELECT 'X' FROM KEY_SURVEYKIND WHERE SVID = @SVID "
    '    Dim sCmd As New SqlCommand(sql, oConn)
    '    TIMS.OpenDbConn(oConn)
    '    With sCmd
    '        .Parameters.Clear()
    '        .Parameters.Add("SVID", SqlDbType.VarChar).Value = SVID
    '        rst.Load(.ExecuteReader())
    '    End With
    '    Return rst
    'End Function

    'Function Get_dtSdS(ByVal SVID As String, ByRef oConn As SqlConnection) As DataTable
    '    Dim rst As New DataTable
    '    Dim sql As String = ""
    '    sql = " SELECT 'X' FROM STUD_SURVEY WHERE SVID =@SVID"
    '    Dim sCmd As New SqlCommand(sql, oConn)
    '    TIMS.OpenDbConn(oConn)
    '    With sCmd
    '        .Parameters.Clear()
    '        .Parameters.Add("SVID", SqlDbType.VarChar).Value = SVID
    '        rst.Load(.ExecuteReader())
    '    End With
    '    Return rst
    'End Function

#End Region

    '儲存
    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        'Dim sql As String
        'Dim dr As DataRow
        Dim ISUSE2 As String = "N"
        If ISUSE.Checked = True Then '是否啟用
            ISUSE2 = "Y"
        End If
        Dim CHKinternalVAL As String = ""
        If chkinternal.Checked Then
            CHKinternalVAL = "Y"
        End If
        IputQName.Value = TIMS.ClearSQM(IputQName.Value)

        Select Case HidMode.Value
            Case "I" '新增
                If IputQName.Value = "" Then  '問卷名稱是空值時
                    Common.MessageBox(Me, "請輸入問卷名稱")
                    Exit Sub
                End If
                If IputQName.Value <> "" Then  '問卷名稱不是空值時
                    Dim sql As String = ""
                    sql = "SELECT 'X' FROM ID_SURVEY WHERE NAME  = '" & IputQName.Value & "'"
                    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
                    If Not dr Is Nothing Then
                        Common.MessageBox(Me, "新增失敗，問卷名稱重複!!")
                        Exit Sub
                    End If
                End If

            Case "E" '修改
                If HidSVID.Value = "" Then
                    Common.MessageBox(Me, "修改失敗，請輸入問卷ID")
                    Exit Sub
                End If
                If IputQName.Value = "" Then '問卷名稱是空值時
                    Common.MessageBox(Me, "請輸入問卷名稱")
                    Exit Sub
                End If
                Dim sql As String = ""
                sql = "Select 'x' from ID_Survey Where Name  = '" & IputQName.Value & "' and SVID != '" & HidSVID.Value & "'"
                Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    Common.MessageBox(Me, "修改失敗，問卷名稱重複!!")
                    Exit Sub
                End If
                sql = "Select 'x' from ID_Survey Where SVID = '" & HidSVID.Value & "'"
                dr = DbAccess.GetOneRow(sql, objconn)
                If dr Is Nothing Then
                    Common.MessageBox(Me, "修改失敗，查無問卷ID")
                    Exit Sub
                End If

        End Select

        Select Case HidMode.Value
            Case cst_Insert ' "I"
                Dim SS As String = ""
                TIMS.SetMyValue(SS, "SurveyName", IputQName.Value)
                TIMS.SetMyValue(SS, "ISUSE2", ISUSE2)
                TIMS.SetMyValue(SS, "CHKinternalVAL", CHKinternalVAL)
                HidSVID.Value = CStr(UpdateInsertSurvey(cst_Insert, SS))
                If Val(HidSVID.Value) = 0 Then
                    Common.MessageBox(Me, "新增失敗!!")
                    Exit Sub
                End If

                Dim sql As String = ""
                sql = ""
                sql &= " INSERT INTO ID_SurveyTypeRel(STID,SVID, ModifyAcct,MODIFYDATE) "
                sql += " values(@STID,@SVID, @ModifyAcct,getdate())"
                Dim iCmd As New SqlCommand(sql, objconn)
                TIMS.OpenDbConn(objconn)
                With iCmd
                    .Parameters.Clear()
                    '問卷種類
                    If ddlSurveyType.SelectedValue <> "" Then
                        .Parameters.Add("STID", SqlDbType.VarChar).Value = ddlSurveyType.SelectedValue
                    Else
                        .Parameters.Add("STID", SqlDbType.VarChar).Value = " "
                    End If
                    .Parameters.Add("SVID", SqlDbType.VarChar).Value = HidSVID.Value
                    'chkinternal
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .ExecuteNonQuery()
                End With
                Common.MessageBox(Me, "新增成功")
                Call dt_search()

            Case cst_Update '"E"
                Call TIMS.OpenDbConn(objconn)

                Dim SS As String = ""
                TIMS.SetMyValue(SS, "SurveyName", IputQName.Value)
                TIMS.SetMyValue(SS, "ISUSE2", ISUSE2)
                TIMS.SetMyValue(SS, "CHKinternalVAL", CHKinternalVAL)
                TIMS.SetMyValue(SS, "ValSVID", HidSVID.Value)
                HidSVID.Value = CStr(UpdateInsertSurvey(cst_Update, SS))
                If Val(HidSVID.Value) = 0 Then
                    Common.MessageBox(Me, "修改失敗!!")
                    Exit Sub
                End If

                Dim sql As String = ""
                If ddlSurveyType.SelectedValue <> "" Then
                    sql = " select * from ID_SurveyTypeRel where SVID=@SVID "
                    Dim sCmd2 As New SqlCommand(sql, objconn)

                    sql = ""
                    sql &= " INSERT INTO ID_SurveyTypeRel (STID,SVID,ModifyAcct,ModifyDate) "
                    sql &= " values (@STID,@SVID,@ModifyAcct,getdate() ) "
                    Dim iCmd2 As New SqlCommand(sql, objconn)

                    sql = ""
                    sql &= " UPDATE ID_SurveyTypeRel "
                    sql &= " SET STID=@STID "
                    sql &= " ,ModifyAcct=@ModifyAcct "
                    sql &= " ,ModifyDate=getdate() "
                    sql &= " WHERE SVID=@SVID "
                    Dim uCmd2 As New SqlCommand(sql, objconn)

                    Dim dt2 As New DataTable
                    With sCmd2
                        .Parameters.Clear()
                        .Parameters.Add("SVID", SqlDbType.VarChar).Value = HidSVID.Value
                        dt2.Load(.ExecuteReader())
                    End With

                    If dt2.Rows.Count = 0 Then
                        With iCmd2
                            .Parameters.Clear()
                            .Parameters.Add("STID", SqlDbType.VarChar).Value = ddlSurveyType.SelectedValue
                            .Parameters.Add("SVID", SqlDbType.VarChar).Value = HidSVID.Value
                            .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                            .ExecuteNonQuery()
                        End With
                    Else
                        With uCmd2
                            .Parameters.Clear()
                            .Parameters.Add("STID", SqlDbType.VarChar).Value = ddlSurveyType.SelectedValue
                            .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                            .Parameters.Add("SVID", SqlDbType.VarChar).Value = HidSVID.Value
                            .ExecuteNonQuery()
                        End With
                    End If
                End If

                Common.MessageBox(Me, "修改成功")
                dt_search()

        End Select

    End Sub

    '回上一頁
    Protected Sub btnReturn1_Click(sender As Object, e As EventArgs) Handles btnReturn1.Click
        dt_search()
    End Sub

    '查詢
    Protected Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Call dt_search()
    End Sub

    '新增
    Protected Sub btnInsert_Click(sender As Object, e As EventArgs) Handles btnInsert.Click

        '新增
        table_F.Visible = False
        table_I.Visible = True
        DataGrid1.Visible = False
        msg.Visible = False
        PageControler1.Visible = False
        IputQName.Value = ""
        Ipt_Name.Value = ""
        ISUSE.Checked = True
        HidMode.Value = "I"
        ddlSurveyType = TIMS.Get_SurveyType(ddlSurveyType, objconn)
        Common.SetListItem(ddlSurveyType, "01") '預設
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class

