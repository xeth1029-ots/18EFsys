Public Class SD_05_036
    Inherits AuthBasePage

    'ADP_DISASTER,ADP_DISASTER2, ADP_DISASTERDEL
    Const cst_SD_05_036c_ZIPCODE As String = "SD_05_036c_ZIPCODE"
    Const cst_dgAct_Update As String = "ActUpdate1"
    Const cst_dgAct_Delete As String = "ActDelete1"

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
        If hid_sessName1.Value <> "" Then Session(hid_sessName1.Value) = Nothing
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            Call ShowList(0)
        End If
    End Sub

    ''' <summary>
    ''' 刪除資料
    ''' </summary>
    ''' <param name="sCmdArg"></param>
    Sub SDeleteData1(ByVal sCmdArg As String)
        Dim ADID As String = TIMS.GetMyValue(sCmdArg, "ADID")
        If ADID = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg3)
            Exit Sub
        End If
        'hid_ADID.Value = ADID

        Dim pmsU1 As New Hashtable From {{"MODIFYACCT", sm.UserInfo.UserID}, {"ADID", Val(ADID)}}
        Dim SqlU1 As String = ""
        SqlU1 &= " UPDATE ADP_DISASTER SET MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE() WHERE ADID=@ADID" & vbCrLf
        DbAccess.ExecuteNonQuery(SqlU1, objconn, pmsU1)

        Dim pmsS1 As New Hashtable From {{"ADID", Val(ADID)}, {"MODIFYACCT", sm.UserInfo.UserID}}
        Dim SqlS1 As String = ""
        SqlS1 &= " INSERT INTO ADP_DISASTERDEL (ADID,CNAME,BEGDATE,ENDDATE,ALARMMSG1,MEMO1,FUNC1,FUNC2,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE)" & vbCrLf
        SqlS1 &= " SELECT ADID,CNAME,BEGDATE,ENDDATE,ALARMMSG1,MEMO1,FUNC1,FUNC2,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE" & vbCrLf
        SqlS1 &= " FROM ADP_DISASTER WHERE ADID=@ADID AND MODIFYACCT=@MODIFYACCT" & vbCrLf
        DbAccess.ExecuteNonQuery(SqlS1, objconn, pmsS1)

        Dim pmsD1 As New Hashtable From {{"ADID", Val(ADID)}, {"MODIFYACCT", sm.UserInfo.UserID}}
        Dim SqlD1 As String = ""
        SqlD1 &= " DELETE ADP_DISASTER WHERE ADID=@ADID AND MODIFYACCT=@MODIFYACCT" & vbCrLf
        DbAccess.ExecuteNonQuery(SqlD1, objconn, pmsD1)

        Call ShowList(1)
        Call ClearData1()
        Call SSearch1()
    End Sub

    '查詢
    Sub SSearch1()
        schCNAME.Text = TIMS.ClearSQM(schCNAME.Text)

        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim pms1 As New Hashtable
        Dim sql As String = ""
        sql &= " select a.ADID ,a.CNAME" & vbCrLf
        sql &= " ,format(a.BEGDATE,'yyyy/MM/dd') BEGDATE" & vbCrLf
        sql &= " ,format(a.ENDDATE,'yyyy/MM/dd') ENDDATE" & vbCrLf
        sql &= " ,a.ALARMMSG1 ,a.MEMO1, a.FUNC1, a.FUNC2" & vbCrLf
        sql &= " ,a.CREATEACCT ,a.CREATEDATE" & vbCrLf
        sql &= " ,a.MODIFYACCT ,a.MODIFYDATE" & vbCrLf
        sql &= " FROM ADP_DISASTER a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If schCNAME.Text <> "" Then
            pms1.Add("CNAME", schCNAME.Text)
            sql &= " and CNAME LIKE '%'+@CNAME+'%'" & vbCrLf
        End If
        sql &= " ORDER BY a.ADID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, pms1)

        If dt.Rows.Count = 0 Then
            Call ShowList(1)
            msg1.Text = TIMS.cst_NODATAMsg1
            Exit Sub
        End If

        Call ShowList(2)
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '查詢1筆，並顯示 依 sCmdArg.ADID
    Sub SShowData1(ByVal sCmdArg As String)
        Dim ADID As String = TIMS.GetMyValue(sCmdArg, "ADID")
        Dim DSTER_SHOW As String = TIMS.GetMyValue(sCmdArg, "SHOW")
        If ADID = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If
        'hid_ADID.Value = ADID

        Dim pms1 As New Hashtable From {{"ADID", Val(ADID)}}
        Dim sql As String = ""
        sql &= " select a.ADID ,a.CNAME" & vbCrLf
        sql &= " ,format(a.BEGDATE,'yyyy/MM/dd') BEGDATE" & vbCrLf
        sql &= " ,format(a.ENDDATE,'yyyy/MM/dd') ENDDATE" & vbCrLf
        sql &= " ,a.ALARMMSG1 ,a.MEMO1, a.FUNC1, a.FUNC2" & vbCrLf
        sql &= " ,a.CREATEACCT ,a.CREATEDATE" & vbCrLf
        sql &= " ,a.MODIFYACCT ,a.MODIFYDATE" & vbCrLf
        sql &= " FROM ADP_DISASTER  a" & vbCrLf
        sql &= " WHERE a.ADID=@ADID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, pms1)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If

        Dim dr As DataRow = dt.Rows(0)
        hid_ADID.Value = Convert.ToString(dr("ADID"))
        CNAME.Text = Convert.ToString(dr("CNAME"))
        BEGDATE.Text = Convert.ToString(dr("BEGDATE"))
        ENDDATE.Text = Convert.ToString(dr("ENDDATE"))
        ALARMMSG1.Text = Convert.ToString(dr("ALARMMSG1"))
        MEMO1.Text = Convert.ToString(dr("MEMO1"))
        Dim tmp_DISASTER2_N As String = ""
        If DSTER_SHOW = "1" Then
            tmp_DISASTER2_N = TIMS.GetDISASTER2_N(hid_ADID.Value, objconn, 1)
        Else
            tmp_DISASTER2_N = TIMS.GetDISASTER2_N(hid_ADID.Value, objconn, 2)
        End If
        Label1.Text = tmp_DISASTER2_N
        '查詢完整受災地區
        Button4.Visible = (tmp_DISASTER2_N.IndexOf("...") > -1)

        CB1_FUNC1.Checked = (Convert.ToString(dr("FUNC1")) = "Y")
        CB1_FUNC2.Checked = (Convert.ToString(dr("FUNC2")) = "Y")

        hid_ZIPCODE.Value = TIMS.GetDISASTER2(hid_ADID.Value, objconn)
        Session(cst_SD_05_036c_ZIPCODE) = hid_ZIPCODE.Value

        If hid_ZIPCODE.Value = "" Then Label1.Text = "查無 重大災害受災地區!!"

        Call ShowList(4)

    End Sub

    '[儲存前]檢核
    Function CheckData1(ByRef sErrmsg As String) As Boolean
        Dim rst As Boolean = True
        sErrmsg = ""

        hid_ADID.Value = TIMS.ClearSQM(hid_ADID.Value)
        CNAME.Text = TIMS.ClearSQM(CNAME.Text)
        BEGDATE.Text = TIMS.ClearSQM(BEGDATE.Text)
        ENDDATE.Text = TIMS.ClearSQM(ENDDATE.Text)
        ALARMMSG1.Text = TIMS.ClearSQM(ALARMMSG1.Text)
        MEMO1.Text = TIMS.ClearSQM(MEMO1.Text)

        If CNAME.Text = "" Then sErrmsg &= "重大災害名稱 為必填欄位!" & vbCrLf
        If BEGDATE.Text = "" Then sErrmsg &= "起始日期 為必填欄位!" & vbCrLf
        If ENDDATE.Text = "" Then sErrmsg &= "結束日期 為必填欄位!" & vbCrLf
        If CB1_FUNC2.Checked AndAlso ALARMMSG1.Text = "" Then sErrmsg &= "系統告警訊息 為必填欄位!" & vbCrLf
        If Not CB1_FUNC1.Checked AndAlso Not CB1_FUNC2.Checked Then sErrmsg &= "使用功能 必須勾選一項!" & vbCrLf
        If sErrmsg <> "" Then Return False

        If Not TIMS.IsDate1(BEGDATE.Text) Then sErrmsg &= "起始日期 必須是正確的日期格式!" & vbCrLf
        If Not TIMS.IsDate1(ENDDATE.Text) Then sErrmsg &= "結束日期 必須是正確的日期格式!" & vbCrLf
        If sErrmsg <> "" Then Return False

        '若時間順序有誤，修正時間順序
        If DateDiff(DateInterval.Second, CDate(BEGDATE.Text), CDate(ENDDATE.Text)) < 0 Then
            Dim tmp_Date As String = BEGDATE.Text
            BEGDATE.Text = ENDDATE.Text
            ENDDATE.Text = tmp_Date
        End If

        Dim sql3 As String = "SELECT 'X' FROM ADP_DISASTER WITH(NOLOCK) WHERE CNAME=@CNAME AND ADID!=@ADID"

        If hid_ADID.Value = "" Then
            '新增
            Dim pms2 As New Hashtable From {{"CNAME", CNAME.Text}}
            Dim sql2 As String = "SELECT 'X' FROM ADP_DISASTER WHERE CNAME=@CNAME"
            Dim dt2 As DataTable = DbAccess.GetDataTable(sql2, objconn, pms2)
            If dt2.Rows.Count > 0 Then
                sErrmsg &= "重大災害名稱 已存在資料庫，請修改其它名稱!" & vbCrLf
            End If
        Else
            '修改
            Dim pms2 As New Hashtable From {{"CNAME", CNAME.Text}, {"ADID", Val(hid_ADID.Value)}}
            Dim sql2 As String = "SELECT 'X' FROM ADP_DISASTER WHERE CNAME=@CNAME AND ADID!=@ADID"
            Dim dt2 As DataTable = DbAccess.GetDataTable(sql2, objconn, pms2)
            If dt2.Rows.Count > 0 Then
                sErrmsg &= "重大災害名稱 已存在資料庫，請修改其它名稱!" & vbCrLf
            End If
        End If
        If sErrmsg <> "" Then Return False

        If hid_sessName1.Value <> "" Then
            hid_ZIPCODE.Value = Convert.ToString(Session(hid_sessName1.Value))
            'Session(hid_sessName1.Value) = Nothing
        End If
        If hid_ZIPCODE.Value = "" Then
            sErrmsg &= "未選擇災害受災地區，資料不完整，無法正確運行." & vbCrLf 'Common.MessageBox(Me, TIMS.cst_NODATAMsg3) 'Common.MessageBox(Me, vMsg)
            Return False 'Exit Sub
        End If

        Dim objA As String() = hid_ZIPCODE.Value.Split(",")
        Dim tmpV As String = ""
        For Each strZIPCODE As String In objA
            strZIPCODE = TIMS.ClearSQM(strZIPCODE)
            If strZIPCODE <> "" Then
                tmpV = "1"
                Exit For
            End If
        Next
        If tmpV = "" Then
            'vMsg = "未選擇災害受災地區，資料不完整，無法正確運行!!" 'Common.MessageBox(Me, TIMS.cst_NODATAMsg3)
            sErrmsg &= "未選擇災害受災地區，資料不完整，無法正確運行.." & vbCrLf ' Common.MessageBox(Me, vMsg)
            Return False 'Exit Sub
        End If

        If sErrmsg <> "" Then Return False
        Return rst
    End Function

    '儲存
    Sub SaveData1()
        If TIMS.ChkSession(sm) Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg8)
            Exit Sub
        End If

        hid_ADID.Value = TIMS.ClearSQM(hid_ADID.Value)
        CNAME.Text = TIMS.ClearSQM(CNAME.Text)
        BEGDATE.Text = TIMS.ClearSQM(BEGDATE.Text)
        ENDDATE.Text = TIMS.ClearSQM(ENDDATE.Text)
        ALARMMSG1.Text = TIMS.ClearSQM(ALARMMSG1.Text)
        MEMO1.Text = TIMS.ClearSQM(MEMO1.Text)
        Dim v_FUNC1 As String = If(CB1_FUNC1.Checked, "Y", "")
        Dim v_FUNC2 As String = If(CB1_FUNC2.Checked, "Y", "")

        Dim sqlins As String = ""
        sqlins &= " INSERT INTO ADP_DISASTER (ADID,CNAME,BEGDATE,ENDDATE,ALARMMSG1,MEMO1,FUNC1,FUNC2,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE)" & vbCrLf
        sqlins &= " VALUES(@ADID,@CNAME,@BEGDATE,@ENDDATE,@ALARMMSG1,@MEMO1,@FUNC1,@FUNC2,@CREATEACCT,GETDATE(),@MODIFYACCT,GETDATE())" & vbCrLf
        Dim iCmd As New SqlCommand(sqlins, objconn)

        Dim sqlupd As String = ""
        sqlupd &= " update ADP_DISASTER" & vbCrLf
        sqlupd &= " set CNAME=@CNAME ,BEGDATE=@BEGDATE ,ENDDATE=@ENDDATE ,ALARMMSG1=@ALARMMSG1" & vbCrLf
        sqlupd &= " ,MEMO1=@MEMO1 ,FUNC1=@FUNC1 ,FUNC2=@FUNC2 ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
        sqlupd &= " WHERE ADID=@ADID" & vbCrLf
        Dim uCmd As New SqlCommand(sqlupd, objconn)

        Call TIMS.OpenDbConn(objconn)

        If hid_ADID.Value = "" Then
            '新增
            Dim iADID As Integer = DbAccess.GetNewId(objconn, "ADP_DISASTER_ADID_SEQ,ADP_DISASTER,ADID")
            With iCmd
                .Parameters.Clear()
                .Parameters.Add("ADID", SqlDbType.Int).Value = iADID
                .Parameters.Add("CNAME", SqlDbType.VarChar).Value = CNAME.Text
                .Parameters.Add("BEGDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(BEGDATE.Text)
                .Parameters.Add("ENDDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(ENDDATE.Text)
                .Parameters.Add("ALARMMSG1", SqlDbType.NVarChar).Value = ALARMMSG1.Text
                .Parameters.Add("MEMO1", SqlDbType.NVarChar).Value = MEMO1.Text
                .Parameters.Add("FUNC1", SqlDbType.VarChar).Value = If(v_FUNC1 <> "", v_FUNC1, Convert.DBNull)
                .Parameters.Add("FUNC2", SqlDbType.VarChar).Value = If(v_FUNC2 <> "", v_FUNC2, Convert.DBNull)
                .Parameters.Add("CREATEACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .ExecuteNonQuery()
                'dt.Load(.ExecuteReader())
            End With
            hid_ADID.Value = iADID
        Else
            '修改
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("CNAME", SqlDbType.VarChar).Value = CNAME.Text
                .Parameters.Add("BEGDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(BEGDATE.Text)
                .Parameters.Add("ENDDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(ENDDATE.Text)
                .Parameters.Add("ALARMMSG1", SqlDbType.NVarChar).Value = ALARMMSG1.Text
                .Parameters.Add("MEMO1", SqlDbType.NVarChar).Value = MEMO1.Text
                .Parameters.Add("FUNC1", SqlDbType.VarChar).Value = If(v_FUNC1 <> "", v_FUNC1, Convert.DBNull)
                .Parameters.Add("FUNC2", SqlDbType.VarChar).Value = If(v_FUNC2 <> "", v_FUNC2, Convert.DBNull)
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("ADID", SqlDbType.Int).Value = Val(hid_ADID.Value)
                .ExecuteNonQuery()
                'dt.Load(.ExecuteReader())
            End With
        End If

        Dim fg_SAVEOK As Boolean = SaveData2(hid_ADID.Value)
        If Not fg_SAVEOK Then Return

        Call ShowList(1)
        Call ClearData1()
        Call SSearch1()
    End Sub

    '儲存2 異常返回 false
    Function SaveData2(bvADID As String) As Boolean
        Dim vMsg As String = ""
        If TIMS.ChkSession(sm) Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg8)
            Return False 'Exit Sub
        End If
        If bvADID = "" Then
            Common.MessageBox(Me, String.Concat(TIMS.cst_NODATAMsg3, ",bvADID =''"))
            Return False 'Exit Sub
        End If
        If hid_sessName1.Value <> "" Then
            hid_ZIPCODE.Value = Convert.ToString(Session(hid_sessName1.Value))
            Session(hid_sessName1.Value) = Nothing
        End If
        If hid_ZIPCODE.Value = "" Then
            vMsg = "未選擇災害受災地區，資料不完整，無法正確運行!" 'Common.MessageBox(Me, TIMS.cst_NODATAMsg3)
            Common.MessageBox(Me, vMsg)
            Return False 'Exit Sub
        End If

        Dim objA As String() = hid_ZIPCODE.Value.Split(",")
        Dim tmpV As String = ""
        For Each strZIPCODE As String In objA
            strZIPCODE = TIMS.ClearSQM(strZIPCODE)
            If strZIPCODE <> "" Then
                tmpV = "1"
                Exit For
            End If
        Next
        If tmpV = "" Then
            vMsg = "未選擇災害受災地區，資料不完整，無法正確運行!!" 'Common.MessageBox(Me, TIMS.cst_NODATAMsg3)
            Common.MessageBox(Me, vMsg)
            Return False 'Exit Sub
        End If

        Try
            Dim sqldel As String = " DELETE ADP_DISASTER2 WHERE ADID=@ADID" & vbCrLf
            Using dCmd As New SqlCommand(sqldel, objconn)
                With dCmd
                    .Parameters.Clear()
                    .Parameters.Add("ADID", SqlDbType.Int).Value = Val(bvADID)
                    .ExecuteNonQuery()
                    'dt.Load(.ExecuteReader())
                End With
            End Using

            For Each strZIPCODE As String In objA
                strZIPCODE = TIMS.ClearSQM(strZIPCODE)
                If strZIPCODE <> "" Then
                    Dim sqlinto As String = " INSERT INTO ADP_DISASTER2(ADID,ZIPCODE,USE1,MODIFYACCT,MODIFYDATE) VALUES (@ADID,@ZIPCODE,'Y',@MODIFYACCT,GETDATE())"
                    Using iCmd As New SqlCommand(sqlinto, objconn)
                        With iCmd
                            .Parameters.Clear()
                            .Parameters.Add("ADID", SqlDbType.Int).Value = Val(bvADID)
                            .Parameters.Add("ZIPCODE", SqlDbType.VarChar).Value = strZIPCODE
                            .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                            .ExecuteNonQuery() 'dt.Load(.ExecuteReader())
                        End With
                    End Using
                End If
            Next

        Catch ex As Exception
            Call TIMS.WriteTraceLog(ex.Message, ex)
            Common.MessageBox(Me, TIMS.cst_NODATAMsg3)
            Return False 'Exit Sub
        End Try
        Return True
    End Function

    '清理欄位
    Sub ClearData1()
        hid_ADID.Value = ""
        CNAME.Text = ""
        BEGDATE.Text = ""
        ENDDATE.Text = ""
        ALARMMSG1.Text = ""
        MEMO1.Text = ""
        CB1_FUNC1.Checked = True
        CB1_FUNC2.Checked = True
        Button4.Visible = False

        hid_ZIPCODE.Value = ""
        Label1.Text = ""
    End Sub

    '顯示依 iType
    Sub ShowList(ByVal iType As Integer)
        msg1.Text = ""
        tbSearch1.Visible = False
        tbDataGrid1.Visible = False
        tbDetail1.Visible = False
        Select Case iType
            Case 0 '初始呼叫
                tbSearch1.Visible = True
            Case 1 '查詢/無查詢結果
                tbSearch1.Visible = True
            Case 2 '顯示查詢結果
                tbSearch1.Visible = True
                tbDataGrid1.Visible = True
            Case 4 '新增一筆/修改一筆
                Button3.Disabled = False '(啟用)
                tbDetail1.Visible = True
        End Select
    End Sub

    '修改/刪除
    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        '修改/刪除
        If e.CommandArgument = "" Then Exit Sub
        Dim sCmdArg As String = e.CommandArgument
        'Dim ADID As String = TIMS.GetMyValue(sCmdArg, "ADID")
        'If ADID = "" Then Exit Sub
        Select Case e.CommandName
            Case cst_dgAct_Delete
                Call SDeleteData1(sCmdArg)

            Case cst_dgAct_Update
                Call ClearData1()
                Call SShowData1(sCmdArg)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim lbtUpdate1 As LinkButton = e.Item.FindControl("lbtUpdate1")
                Dim lbtDelete1 As LinkButton = e.Item.FindControl("lbtDelete1")
                lbtUpdate1.CommandName = cst_dgAct_Update
                lbtDelete1.CommandName = cst_dgAct_Delete

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "ADID", Convert.ToString(drv("ADID")))
                lbtUpdate1.CommandArgument = sCmdArg
                lbtDelete1.CommandArgument = sCmdArg
                lbtDelete1.Attributes("onclick") = "return confirm('您確定要刪除此筆資料?');"

                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + sender.CurrentPageIndex * sender.PageSize

        End Select
    End Sub

    '查詢
    Protected Sub BtnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
        Call SSearch1()
    End Sub

    '新增按鈕
    Protected Sub BtnInsert1_Click(sender As Object, e As EventArgs) Handles btnInsert1.Click
        Call ClearData1()
        Call ShowList(4)
    End Sub

    '儲存
    Protected Sub BtnSaveData1_Click(sender As Object, e As EventArgs) Handles btnSaveData1.Click
        '[儲存前]檢核
        Dim sErrmsg As String = ""
        Call CheckData1(sErrmsg)
        If sErrmsg <> "" Then
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If

        Call SaveData1()
    End Sub

    '回上頁
    Protected Sub BtnBackup1_Click(sender As Object, e As EventArgs) Handles btnBackup1.Click
        Call ShowList(1)
        Call SSearch1()
    End Sub

    Protected Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        hid_ADID.Value = TIMS.ClearSQM(hid_ADID.Value)
        Dim sCmdArg As String = ""
        TIMS.SetMyValue(sCmdArg, "ADID", hid_ADID.Value)
        TIMS.SetMyValue(sCmdArg, "SHOW", "1")
        Call SShowData1(sCmdArg)
    End Sub
End Class
