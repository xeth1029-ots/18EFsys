Partial Class TrainJob
    Inherits AuthBasePage

    'Key_TrainType
    'TrainJob.aspx? / openTrain2
    'select * from Key_TrainType where tmid =197
    'select * from Key_TrainType where parent=197
    'Dim sql As String = ""
    Const cst_TPlanIDtype_TIMS As String = "1"
    Const cst_TPlanIDtype_TIMS28 As String = "2"
    Const cst_TPlanIDtype_TIMS28_18 As String = "3" '2018以後

    Dim dtG As DataTable = Nothing
    Dim ff As String = ""
    Dim fsort As String = ""
    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        iPYNum = TIMS.sUtl_GetPYNum(Me)
        Call cCreat0()

        'If Not IsPostBack Then Call cCreat1()  'edit，by:20181026
        '20181026 依照107增修需求,"主要職類"裡的"支用標準"直接設定為"職類課程分類" ---------------------- start
        If Not IsPostBack Then
            Call cCreat1()
            If busTD.InnerHtml = "支用標準" Then
                Dim t_Index As Integer = bus.Items.IndexOf(bus.Items.FindByText("職類課程分類"))
                If t_Index > -1 Then
                    bus.SelectedIndex = t_Index
                    Chg_BusSelChanged()
                End If
            End If
        End If
        '-------------------------------------------------------------------------------------- end
    End Sub

    ''' <summary> 每次執行 </summary>
    Sub cCreat0()
        Hid_TPlanIDtype123.Value = cst_TPlanIDtype_TIMS '"1" '1@TIMS 2@TIMS28 3@TIMS-28(2018)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Hid_TPlanIDtype123.Value = cst_TPlanIDtype_TIMS28 '"2"
            If iPYNum >= 3 Then Hid_TPlanIDtype123.Value = cst_TPlanIDtype_TIMS28_18 '"3"
        End If
        Dim sql As String = ""
        'vsRequest_type= Request("type") '沒有／數字
        Dim rqfield As String = TIMS.ClearSQM(Request("field")) '不定
        Dim RqTMID As String = TIMS.ClearSQM(Request("TMID")) '沒有／文數字
        'fieldname.Value = TIMS.ClearSQM(rqfield)'
        If TIMS.CheckInput(rqfield) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg2)
            Exit Sub
        End If
        If TIMS.CheckInput(RqTMID) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg2)
            Exit Sub
        End If
        fieldname.Value = TIMS.ClearSQM(rqfield)
        Select Case Hid_TPlanIDtype123.Value
            Case cst_TPlanIDtype_TIMS
                'TIMS計畫 非產投
                'sqlstr = "select *,'['+TrainID+']'+TrainName as NewTrainName from Key_TrainType WHERE (BusID<>'G' or BusID IS NULL) order by TMID"
                '因造成統計誤，故有關訓練職類,把其他類拿掉,日後不顯示其他類!! amu 2008/08
                'Dim sql As String = ""
                'sql = "" & vbCrLf
                'sql &= " SELECT a.* ,'[' + a.TrainID + ']' + a.TrainName NewTrainName ,'[' + dbo.LPAD(a.JobID,2,'0') + ']' + a.JobName NewJobName ,dbo.LPAD(a.JobID,2,'0') LLJOBID " & vbCrLf
                'sql &= " FROM Key_TrainType A " & vbCrLf
                'sql &= " WHERE (a.BusID < 'G' OR a.BusID IS NULL) AND a.TMID !=6 AND a.TMID < 555 " & vbCrLf
                'sql &= " ORDER BY a.TMID" & vbCrLf
                sql = ""
                sql &= " SELECT a.TMID,a.BUSID,a.BUSNAME,a.JOBID,a.JOBNAME,a.TRAINID,a.TRAINNAME" & vbCrLf
                sql &= " ,a.LEVELS,a.PARENT,a.MEMO,a.GCID,a.GCID2,a.GCID3" & vbCrLf
                sql &= " ,a.NewTrainName,a.NewJobName,a.LLJOBID" & vbCrLf
                sql &= " FROM dbo.VIEW_TRAINTYPE3 a" & vbCrLf
                sql &= " ORDER BY a.TMID" & vbCrLf

            Case cst_TPlanIDtype_TIMS28
                Page.RegisterStartupScript("loading", "<script>document.title='請選擇業別';</script>")
                '97年產業人才投資方案改為兩欄 (計劃別/訓練業別) 供選擇
                '因造成統計誤，故有關訓練職類,把其他類拿掉,日後不顯示其他類!! amu 2008/08
                '2015停用 TMID:211 醫療保健與照護產業 BY 201408 AMU
                '2015增加 TMID:554 JOBID:99:其他(產投)
                busTD.InnerHtml = "計劃別"
                jobTD.InnerHtml = "訓練業別"
                trainTR.Visible = False
                'Dim sql As String = ""
                sql = ""
                sql &= " SELECT a.TMID,a.BUSID,a.BUSNAME,a.JOBID,a.JOBNAME,a.TRAINID,a.TRAINNAME" & vbCrLf
                sql &= " ,a.LEVELS,a.PARENT,a.MEMO,a.GCID,a.GCID2,a.GCID3" & vbCrLf
                sql &= " ,case when a.TrainID is not null then concat('[',a.TrainID,']',a.TrainName) end NewTrainName" & vbCrLf
                sql &= " ,case when a.JobID is not null then concat('[', dbo.LPAD(a.JobID,2,'0'),']',a.JobName) end NewJobName" & vbCrLf
                sql &= " ,dbo.LPAD(a.JobID,2,'0') LLJOBID " & vbCrLf
                sql &= " FROM dbo.KEY_TRAINTYPE a" & vbCrLf
                sql &= " WHERE (A.BusID='G' OR A.PARENT=197) " & vbCrLf
                If sm.UserInfo.Years >= "2015" Then sql &= " AND a.TMID!=211 " & vbCrLf '停用:醫療保健與照護產業
                Select Case Convert.ToString(sm.UserInfo.LID) '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
                    Case "0", "1"
                    Case Else '"2"
                        sql &= " AND a.TMID!=554 " & vbCrLf '停用:其他(產投)
                End Select
                sql &= " ORDER BY a.TMID " & vbCrLf
            Case cst_TPlanIDtype_TIMS28_18 '2018以後
                Page.RegisterStartupScript("loading", "<script>document.title='支用標準';</script>")
                busTD.InnerHtml = "支用標準"
                jobTD.InnerHtml = "職類課程"
                trainTD.InnerHtml = "業別" '訓練業別
                'train.AutoPostBack = True
                'Dim sql As String = ""
                sql = "" & vbCrLf
                sql &= " SELECT a.TMID,a.BUSID,a.BUSNAME,a.JOBID,a.JOBNAME,a.TRAINID,a.TRAINNAME" & vbCrLf
                sql &= " ,a.LEVELS,a.PARENT,a.MEMO,a.GCID,a.GCID2,a.GCID3" & vbCrLf
                sql &= " ,case when a.TrainName is not null then concat('[',g.GCODE2,']',a.TrainName) end NewTrainName" & vbCrLf
                sql &= " ,case when a.JobName is not null then concat('[',g.GCODE31,']',a.JobName) end NewJobName" & vbCrLf
                sql &= " ,dbo.LPAD(a.JobID,2,'0') LLJOBID" & vbCrLf
                sql &= " FROM dbo.KEY_TRAINTYPE a" & vbCrLf
                sql &= " LEFT JOIN dbo.V_GOVCLASSCAST3 g on g.GCID3=a.GCID3" & vbCrLf
                sql &= " WHERE (a.BusID='H' OR (a.TMID>=600 AND a.TMID<=754))" & vbCrLf
                sql &= " ORDER BY a.TMID " & vbCrLf
        End Select
        If sql = "" Then Return

        dtG = New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            dtG.Load(.ExecuteReader())
        End With
    End Sub

    ''' <summary>
    ''' 只有第1次才執行
    ''' </summary>
    Sub cCreat1()
        Dim rqfield As String = TIMS.ClearSQM(Request("field")) '不定
        Dim RqTMID As String = TIMS.ClearSQM(Request("TMID")) '沒有／文數字
        Dim sql As String = ""
        ff = "levels='0'"
        fsort = "BusID"
        With Me.bus
            .Items.Clear()
            For Each dr As DataRow In dtG.Select(ff, fsort)
                .Items.Add(New ListItem(dr("BusName"), dr("TMID")))
            Next
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
        Me.job.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        If trainTR.Visible Then Me.train.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        Select Case Hid_TPlanIDtype123.Value
            Case cst_TPlanIDtype_TIMS
                If RqTMID = "" Then
                    '非產業人才投資方案-可以使用此記憶功能，因為規則不同
                    Dim dt As DataTable = TIMS.GetCookieTable(Me, objconn)
                    ff = "ItemName='TrainJob_bus'"
                    If dt.Select(ff).Length <> 0 Then
                        Common.SetListItem(bus, dt.Select(ff)(0)("ItemValue"))
                        Chg_BusSelChanged() ' bus_SelectedIndexChanged(sender, e)
                        ff = "ItemName='TrainJob_job'"
                        If dt.Select(ff).Length <> 0 Then '★增加判斷是否有值
                            Common.SetListItem(job, dt.Select(ff)(0)("ItemValue"))
                            Chg_JobSelChanged() 'job_SelectedIndexChanged(sender, e)
                        End If
                        ff = "ItemName='TrainJob_train'"
                        If dt.Select(ff).Length <> 0 Then
                            If Convert.ToString(dt.Select(ff)(0)("ItemValue")) <> "" Then
                                If trainTR.Visible Then Common.SetListItem(train, dt.Select(ff)(0)("ItemValue"))
                            End If
                        End If
                    End If
                End If
            Case cst_TPlanIDtype_TIMS28_18
                train.AutoPostBack = True
        End Select
        If RqTMID <> "" Then
            Dim TMID1 As Integer = 0
            Dim TMID2 As Integer = 0
            Dim TMID3 As Integer = Val(RqTMID)
            sql = " SELECT Parent FROM Key_TrainType WHERE TMID = '" & Val(TMID3) & "'" 'Request("TMID")
            TMID2 = DbAccess.ExecuteScalar(sql, objconn)
            sql = " SELECT Parent FROM Key_TrainType WHERE TMID = '" & Val(TMID2) & "'"
            TMID1 = DbAccess.ExecuteScalar(sql, objconn)
            'dr = DbAccess.GetOneRow(sql, objConn)
            Common.SetListItem(bus, TMID1)
            Chg_BusSelChanged() 'bus_SelectedIndexChanged(sender, e)
            Common.SetListItem(job, TMID2)
            Chg_JobSelChanged() 'job_SelectedIndexChanged(sender, e)
            If trainTR.Visible Then
                Common.SetListItem(train, TMID3)
                If train.AutoPostBack Then Chg_trainSelChanged()
            End If
        End If
    End Sub

    Sub Chg_BusSelChanged()
        'Dim mydv2 As New DataView(objdataset.Tables(0))
        Dim mydv2 As New DataView(dtG)
        If Me.bus.SelectedValue <> "" Then
            mydv2.RowFilter = "levels='1' and [parent]='" & Me.bus.SelectedValue & "'"
            mydv2.Sort = "LLJOBID"
            Me.job.DataSource = mydv2
            'Me.job.DataTextField = "JobName"
            Me.job.DataTextField = "NewJobName"
            Me.job.DataValueField = "TMID"
            Me.job.DataBind()
            Me.job.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            Me.train.Items.Clear()
        Else
            Me.job.Items.Clear()
            Me.job.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            Me.train.Items.Clear()
            Me.train.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End If
    End Sub

    Sub Chg_JobSelChanged()
        'Dim mydv3 As New DataView(objdataset.Tables(0))
        Dim mydv3 As New DataView(dtG)
        If Me.job.SelectedValue <> "" Then
            mydv3.RowFilter = "levels='2' and [parent]='" & Me.job.SelectedValue & "'"
            mydv3.Sort = "TrainID"
            Me.train.DataSource = mydv3
            Me.train.DataTextField = "NewTrainName"
            Me.train.DataValueField = "TMID"
            Me.train.DataBind()
            Me.train.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        Else
            Me.train.Items.Clear()
            Me.train.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End If
    End Sub

    Sub Chg_trainSelChanged()
        Dim v_train As String = TIMS.GetListValue(train)
        If v_train = "" Then Exit Sub

        Dim ff3 As String = String.Concat("tmid=", v_train)
        If dtG.Select(ff3).Length > 0 Then Hid_PERC100.Value = TIMS.Get_PERC100(v_train, objconn)
    End Sub

    ''' <summary>
    ''' 行業別。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub bus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bus.SelectedIndexChanged
        Chg_BusSelChanged()
    End Sub

    ''' <summary>
    ''' 職業分類。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub job_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles job.SelectedIndexChanged
        Chg_JobSelChanged()
    End Sub

    Protected Sub train_SelectedIndexChanged(sender As Object, e As EventArgs) Handles train.SelectedIndexChanged
        Chg_trainSelChanged()
    End Sub

    ''' <summary>
    ''' 送出
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.ServerClick
        Select Case Hid_TPlanIDtype123.Value
            Case cst_TPlanIDtype_TIMS
                '非產業人才投資方案-可以使用此記憶功能，因為規則不同
                If trainTR.Visible Then
                    Dim dt As DataTable = Nothing
                    Dim da As SqlDataAdapter = Nothing
                    TIMS.InsertCookieTable(Me, dt, da, "TrainJob_bus", bus.SelectedValue, False, objconn)
                    TIMS.InsertCookieTable(Me, dt, da, "TrainJob_job", job.SelectedValue, False, objconn)
                    TIMS.InsertCookieTable(Me, dt, da, "TrainJob_train", train.SelectedValue, True, objconn)
                End If
        End Select
        Common.RespWrite(Me, "<script>window.close();</script>")
    End Sub
End Class