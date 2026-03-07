Partial Class CP_01_ch
    Inherits AuthBasePage

    'Dim dt As DataTable
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 1)
        TIMS.ChkSession(Me, 1, sm)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        'Request("AcctPlan")=1              表示可以跨計畫選擇
        'Request("special")=1               增加回傳開訓跟結訓日期
        'Request("special")=2               母頁Submit
        'Request("special")=3               SD_05_005.aspx專用
        'Request("sort")=3                  表示回傳到TMID3.value

        search_but.Attributes("onclick") = "javascript:return chkdata();"

        'Dim sSql As String
        'sSql = "select * from Key_TrainType WITH(NOLOCK)"
        'dtTrainType = DbAccess.GetDataTable(sSql, objconn)

        If Not IsPostBack Then
            msg.Text = ""
            Table2.Visible = False
            HourRan = TIMS.GET_HOURRAN(HourRan, objconn, sm)

            Dim dt As DataTable = TIMS.GetCookieTable(Me, objconn)

            TB_career_id.Text = TIMS.GetCookieItemValue(dt, "SD_TB_career_id")
            trainValue.Value = TIMS.GetCookieItemValue(dt, "SD_trainValue") 'dt.Select("ItemName='SD_trainValue'")(0)("ItemValue").ToString.Trim
            Common.SetListItem(HourRan, TIMS.GetCookieItemValue(dt, "SD_HourRan"))
            CyclType.Text = TIMS.GetCookieItemValue(dt, "SD_CyclType") 'dt.Select("ItemName='SD_CyclType'")(0)("ItemValue").ToString.Trim
            ClassID.Text = TIMS.GetCookieItemValue(dt, "SD_ClassID") 'dt.Select("ItemName='SD_ClassID'")(0)("ItemValue").ToString.Trim
            Common.SetListItem(ClassRound, TIMS.GetCookieItemValue(dt, "SD_ClassRound"))
            'Call search_but_Click(sender, e)
            If TB_career_id.Text <> "" AndAlso trainValue.Value <> "" Then Call Search1()

        End If
    End Sub

    Sub save_search_Cookie()
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        TIMS.InsertCookieTable(Me, dt, da, "SD_TB_career_id", TB_career_id.Text, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_trainValue", trainValue.Value, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_HourRan", HourRan.SelectedValue, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_CyclType", CyclType.Text, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_ClassID", ClassID.Text, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_ClassRound", ClassRound.SelectedValue, True, objconn)
    End Sub

    Sub Search1()
        Call save_search_Cookie()

        Dim o_parms As New Hashtable
        Dim i_parms As New Hashtable
        i_parms.Add("type1", "2")
        Dim sql As String = Utl_SearchSql(i_parms, o_parms)

        Dim dt2 As DataTable
        dt2 = DbAccess.GetDataTable(sql, objconn, o_parms)

        msg.Text = "查無資料!!"
        Table2.Visible = False
        BTN_send.Visible = False
        If dt2.Rows.Count > 0 Then
            msg.Text = ""
            Table2.Visible = True
            BTN_send.Visible = True

            PageControler1.PageDataTable = dt2
            PageControler1.PrimaryKey = "OCID"
            PageControler1.Sort = "ClassID,CyclType"
            PageControler1.ControlerLoad()
        End If
    End Sub

    '查詢鈕
    Private Sub search_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search_but.Click
        Call Search1()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Const cst_cells_STDate As Integer = 4

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                '開訓日期
                e.Item.Cells(cst_cells_STDate).Text = TIMS.Cdate17(drv("STDate"))
        End Select

    End Sub

    Sub search_data1()
        Dim rqclass1 As String = TIMS.ClearSQM(Request("class1"))
        If rqclass1 = "" Then
            Common.MessageBox(Me, "請先勾選班級!")
            Exit Sub
        End If

        Dim drCC As DataRow = TIMS.GetOCIDDate(rqclass1, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "查無資料!")
            Exit Sub
        End If

        Dim o_parms As New Hashtable
        Dim i_parms As New Hashtable
        i_parms.Clear()
        i_parms.Add("type1", "1")
        i_parms.Add("OCID1", rqclass1)
        Dim sql As String = Utl_SearchSql(i_parms, o_parms)
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, o_parms)
        If dr Is Nothing Then
            Common.MessageBox(Me, "查無資料!")
            Exit Sub
        End If

        Call Utl_RespWrite(dr)
        Call save_Cookie(dr)
    End Sub

    ''' <summary>
    ''' 組合sql type1:1/2
    ''' </summary>
    ''' <param name="i_parms"></param>
    ''' <param name="o_parms"></param>
    ''' <returns></returns>
    Function Utl_SearchSql(ByRef i_parms As Hashtable, ByRef o_parms As Hashtable) As String
        'Dim rqclass1 As String = TIMS.ClearSQM(Request("class1"))
        o_parms = New Hashtable
        Dim type1 As String = TIMS.GetMyValue2(i_parms, "type1")

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT cc.OCID" & vbCrLf
        sql &= " ,cc.CLASSCNAME" & vbCrLf
        sql &= " ,cc.CYCLTYPE" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME2" & vbCrLf

        sql &= " ,concat(cc.Years,'0',ISNULL(ic.ClassID2,ic.ClassID),cc.CyclType) ClassID" & vbCrLf
        sql &= " ,format(cc.SENTERDATE,'yyyy/MM/dd HH:mm:ss') SENTERDATE" & vbCrLf
        sql &= " ,format(cc.FENTERDATE,'yyyy/MM/dd HH:mm:ss') FENTERDATE" & vbCrLf
        sql &= " ,format(cc.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,format(cc.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        sql &= " ,cc.TMID" & vbCrLf
        sql &= " ,cc.RID" & vbCrLf
        sql &= " ,cc.IsApplic" & vbCrLf
        sql &= " ,e.HOURRANNAME " & vbCrLf '訓練時段

        sql &= " ,CASE WHEN cc.IsApplic='Y' THEN '可挑選志願' ELSE '不可挑選志願' END IsApplic2 " & vbCrLf
        sql &= " ,oo.OrgName" & vbCrLf
        sql &= " ,CASE when tt.JobID is null then tt.TrainID else tt.JobID end TrainID" & vbCrLf
        sql &= " ,CASE when tt.JobID is null then tt.trainName else tt.JobName end TrainName" & vbCrLf
        sql &= " ,CASE when tt.JobID is null then concat('[',tt.TrainID,']',tt.trainName) else concat('[',tt.JobID,']',tt.JobName) end TrainName2" & vbCrLf
        '計算開訓人數
        'sql &= " ,dbo.FN_GET_STDCNT(cc.OCID,1) STUDCOUNT" & vbCrLf

        sql &= " FROM CLASS_CLASSINFO cc" & vbCrLf
        sql &= " JOIN KEY_TRAINTYPE tt on tt.TMID=cc.TMID" & vbCrLf
        sql &= " JOIN AUTH_RELSHIP rr on rr.RID=cc.RID" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo on oo.OrgID=rr.OrgID" & vbCrLf
        sql &= " JOIN ID_Class ic on ic.CLSID =cc.CLSID " & vbCrLf
        sql &= " JOIN ID_PLAN ip on ip.PLANID =cc.PLANID " & vbCrLf
        sql &= " LEFT JOIN dbo.KEY_HOURRAN e ON e.HRID = cc.TPeriod" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        'where IsSuccess='Y' and NotOpen='N'
        sql &= " AND cc.IsSuccess='Y'" & vbCrLf
        sql &= " AND cc.NotOpen='N'" & vbCrLf
        Select Case type1
            Case "1" '查詢單筆資料
                Dim OCID1 As String = TIMS.GetMyValue2(i_parms, "OCID1")
                sql &= " AND cc.OCID=@OCID" & vbCrLf
                o_parms.Add("OCID", OCID1)

            Case "2" '查詢多筆資料
                trainValue.Value = TIMS.ClearSQM(trainValue.Value)
                Dim V_HourRan As String = TIMS.GetListValue(HourRan) '.SelectedValue)
                Dim rq_RID As String = TIMS.ClearSQM(Request("RID"))
                CyclType.Text = TIMS.ClearSQM(CyclType.Text)
                ClassID.Text = TIMS.ClearSQM(ClassID.Text)

                sql &= " AND ip.TPlanID=@TPlanID" & vbCrLf
                o_parms.Add("TPlanID", sm.UserInfo.TPlanID)
                sql &= " AND ip.Years=@Years" & vbCrLf
                o_parms.Add("Years", sm.UserInfo.Years)

                rq_RID = If(rq_RID <> "", rq_RID, sm.UserInfo.RID)
                sql &= " and cc.RID=@RID"
                o_parms.Add("RID", rq_RID)

                Select Case sm.UserInfo.LID
                    Case 0
                    Case 1
                        Dim rq_AcctPlan As String = TIMS.ClearSQM(Request("AcctPlan"))
                        sql &= " AND ip.DistID=@DistID" & vbCrLf
                        o_parms.Add("DistID", sm.UserInfo.DistID)
                        If rq_AcctPlan <> "1" Then
                            sql &= " AND ip.PLANID=@PLANID" & vbCrLf
                            o_parms.Add("PLANID", sm.UserInfo.PlanID)
                        End If
                    Case Else
                        sql &= " AND ip.PLANID=@PLANID" & vbCrLf
                        o_parms.Add("PLANID", sm.UserInfo.PlanID)
                End Select

                '班級範圍
                Select Case ClassRound.SelectedIndex
                    Case 0 '開訓二週前
                        sql &= " AND DATEDIFF(DAY,GETDATE(),cc.STDate) <= (2*7) AND DATEDIFF(DAY,GETDATE(),cc.STDate) >0" & vbCrLf
                    Case 1 '已開訓('己開訓改成除排己經結訓的班級)
                        sql &= " AND DATEDIFF(DAY,cc.STDate,GETDATE()) >=0 AND DATEDIFF(DAY,cc.FTDate,GETDATE()) < 0" & vbCrLf
                    Case 2 '已結訓
                        sql &= " AND DATEDIFF(DAY,cc.FTDate,GETDATE()) >=0" & vbCrLf
                    Case 3 '未開訓
                        sql &= " AND DATEDIFF(DAY,GETDATE(),cc.STDate) > 0 " & vbCrLf
                    Case 4 '全部
                    Case Else '異常
                        sql &= " AND 1<>1 " & vbCrLf
                End Select

                If trainValue.Value <> "" Then
                    sql &= " And cc.TMID=@TMID" & vbCrLf
                    o_parms.Add("TMID", trainValue.Value)
                    'TMIDStr = " and cc.TMID='" & trainValue.Value & "'"
                End If
                If V_HourRan <> "" Then
                    sql &= " and cc.TPeriod=@TPeriod" & vbCrLf
                    o_parms.Add("TPeriod", V_HourRan)
                End If

                Dim v_CyclType As String = ""
                If CyclType.Text <> "" Then
                    If IsNumeric(CyclType.Text) Then
                        If Int(CyclType.Text) < 10 Then
                            v_CyclType = "0" & CStr(Int(CyclType.Text))
                        Else
                            v_CyclType = CyclType.Text
                        End If
                    End If
                End If

                If v_CyclType <> "" Then
                    sql &= " and cc.CyclType=@CyclType" & vbCrLf
                    o_parms.Add("CyclType", v_CyclType)
                End If

                If ClassID.Text <> "" Then
                    'If ClassID.Text.Length = 9 Then
                    '    ClassID.Text = Mid(ClassID.Text, 4, 4)
                    'End If
                    sql &= " and concat(cc.Years,'0',ISNULL(ic.ClassID2,ic.ClassID),cc.CyclType) Like '%" & ClassID.Text & "%'" & vbCrLf
                End If

                '班級名稱
                ClassCName.Text = TIMS.ClearSQM(ClassCName.Text)
                If ClassCName.Text <> "" Then sql &= " AND cc.ClassCName LIKE '%" & ClassCName.Text & "%' " & vbCrLf

        End Select

        Return sql
    End Function

    ''' <summary>
    ''' 回傳使用
    ''' </summary>
    ''' <param name="dr"></param>
    Sub Utl_RespWrite(ByRef dr As DataRow)
        If dr Is Nothing Then Return
        Dim className As String = TIMS.GET_CLASSNAME(Convert.ToString(dr("ClassCName")), Convert.ToString(dr("CyclType")))
        Dim r_sort As String = TIMS.ClearSQM(Request("sort"))

        Common.RespWrite(Me, "<script language=javascript>")
        Common.RespWrite(Me, "function returnNum(){")
        If r_sort = "" Then
            Common.RespWrite(Me, "window.opener.document.form1.TMID1.value='[" & dr("TrainID") & "]" & dr("TrainName") & "';")
            Common.RespWrite(Me, "window.opener.document.form1.TMIDValue1.value='" & dr("TMID") & "';")
        Else
            Common.RespWrite(Me, "window.opener.document.form1.TMID" & r_sort & ".value='[" & dr("TrainID") & "]" & dr("TrainName") & "';")
            Common.RespWrite(Me, "window.opener.document.form1.TMIDValue" & r_sort & ".value='" & dr("TMID") & "';")
        End If

        If r_sort = "" Then
            Common.RespWrite(Me, "window.opener.document.form1.OCID1.value='" & className & "';")
            Common.RespWrite(Me, "window.opener.document.form1.OCIDValue1.value='" & dr("OCID") & "';")
        Else
            Common.RespWrite(Me, "window.opener.document.form1.OCID" & r_sort & ".value='" & className & "';")
            Common.RespWrite(Me, "window.opener.document.form1.OCIDValue" & r_sort & ".value='" & dr("OCID") & "';")
        End If

        Select Case Request("special")
            Case 1 '增加回傳開訓與結訓日期
                Common.RespWrite(Me, "window.opener.document.form1.SDate.value='" & dr("STDate") & "';")
                Common.RespWrite(Me, "window.opener.document.form1.FDate.value='" & dr("FTDate") & "';")
            Case 2
                Common.RespWrite(Me, "window.opener.document.form1.submit();")
            Case 3
                'SD_05_005.aspx專用
                Common.RespWrite(Me, "window.opener.document.form1.Button4.disabled=false;")
                Common.RespWrite(Me, "window.opener.document.form1.Button6.disabled=false;")
                Common.RespWrite(Me, "window.opener.document.form1.Usered.value='1';")
        End Select

        Common.RespWrite(Me, "window.close();")
        Common.RespWrite(Me, "}")
        Common.RespWrite(Me, "returnNum();")
        Common.RespWrite(Me, "</script>")
    End Sub

    ''' <summary>
    ''' 儲存查詢條件 
    ''' </summary>
    ''' <param name="dr"></param>
    Sub save_Cookie(ByRef dr As DataRow)
        If dr Is Nothing Then Return
        Dim className As String = TIMS.GET_CLASSNAME(Convert.ToString(dr("ClassCName")), Convert.ToString(dr("CyclType")))
        '存入暫存資料
        Dim da As SqlDataAdapter = Nothing
        'Dim drTemp As DataRow
        Dim dt As DataTable = TIMS.GetCookieTable(Me, da, objconn)
        For i As Integer = 1 To 10
            If dt.Select("ItemName='Temp_OCID" & i & "'").Length = 0 Then
                Dim bInsertFlas As Boolean = True
                For j As Integer = 1 To 10
                    If dt.Select("ItemName='Temp_OCID" & j & "' and ItemValue='" & dr("OCID") & "'").Length <> 0 Then
                        bInsertFlas = False
                        Exit For
                    End If
                Next
                If bInsertFlas Then
                    TIMS.InsertCookieTable(Me, dt, da, "Temp_OCID" & i, dr("OCID"), False, objconn)
                    TIMS.InsertCookieTable(Me, dt, da, "Temp_ClassName" & i, className, False, objconn)
                    TIMS.InsertCookieTable(Me, dt, da, "Temp_TMID" & i, dr("TMID"), False, objconn)
                    TIMS.InsertCookieTable(Me, dt, da, "Temp_TrainName" & i, "[" & dr("TrainID") & "]" & dr("TrainName"), False, objconn)
                    TIMS.InsertCookieTable(Me, dt, da, "Temp_RID" & i, dr("RID"), False, objconn)
                    TIMS.InsertCookieTable(Me, dt, da, "Temp_OrgName" & i, dr("OrgName"), True, objconn)
                End If
                Exit For
            Else
                If i = 10 Then
                    For j As Integer = 1 To 9
                        TIMS.SetCookieItemValue(dt, "Temp_OCID", j)
                        TIMS.SetCookieItemValue(dt, "Temp_ClassName", j)
                        TIMS.SetCookieItemValue(dt, "Temp_TMID", j)
                        TIMS.SetCookieItemValue(dt, "Temp_TrainName", j)
                        TIMS.SetCookieItemValue(dt, "Temp_RID", j)
                        TIMS.SetCookieItemValue(dt, "Temp_OrgName", j)
                    Next

                    Dim InsertFlag As Boolean = True
                    For j As Integer = 1 To 10
                        If dt.Select(String.Format("ItemName='Temp_OCID{0}' and ItemValue='{1}'", j, dr("OCID"))).Length <> 0 Then
                            InsertFlag = False
                            Exit For
                        End If
                    Next
                    If InsertFlag Then
                        TIMS.InsertCookieTable(Me, dt, da, "Temp_OCID" & i, dr("OCID"), False, objconn)
                        TIMS.InsertCookieTable(Me, dt, da, "Temp_ClassName" & i, className, False, objconn)
                        TIMS.InsertCookieTable(Me, dt, da, "Temp_TMID" & i, dr("TMID"), False, objconn)
                        TIMS.InsertCookieTable(Me, dt, da, "Temp_TrainName" & i, "[" & dr("TrainID") & "]" & dr("TrainName"), False, objconn)
                        TIMS.InsertCookieTable(Me, dt, da, "Temp_RID" & i, dr("RID"), False, objconn)
                        TIMS.InsertCookieTable(Me, dt, da, "Temp_OrgName" & i, dr("OrgName"), True, objconn)
                    End If
                End If
            End If
        Next

    End Sub

    '送出鈕
    Protected Sub BTN_send_Click(sender As Object, e As EventArgs) Handles BTN_send.Click
        search_data1()
    End Sub
End Class
