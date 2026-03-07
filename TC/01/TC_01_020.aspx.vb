Partial Class TC_01_020
    Inherits AuthBasePage

    Sub sUtl_PageInit1()
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        'InitializeComponent()
        Dim strTables As String = "PLAN_PLANINFO"
        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS(strTables, objconn)
        If dt.Rows.Count = 0 Then Exit Sub
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTNAME", TB_ContactName) '聯絡人
        'Call TIMS.sUtl_SetMaxLen(dt, "CONTACTPHONE", TB_ContactPhone) '電話
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTEMAIL", TB_ContactEmail) '電子郵件
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTFAX", TB_ContactFax) '傳真
    End Sub

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        Call sUtl_PageInit1()
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            cCreate1()
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        Select Case sm.UserInfo.LID
            Case 2
                Org.Visible = False
                center.Enabled = False
        End Select
    End Sub

    Sub cCreate1()
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)'AppStage = TIMS.Get_AppStage(AppStage)
        If tr_AppStage_TP28.Visible Then
            AppStage = If(sm.UserInfo.Years >= 2018, TIMS.Get_APPSTAGE2(AppStage), TIMS.Get_AppStage(AppStage))
        End If

    End Sub

    Sub sSearch1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        Dim sDistID As String = If(RIDValue.Value <> "", TIMS.Get_DistID_RID(RIDValue.Value, objconn), TIMS.Get_DistID_RID(sm.UserInfo.RID, objconn))
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        center.Text = TIMS.ClearSQM(center.Text)
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        CyclType.Text = TIMS.ClearSQM(CyclType.Text)

        Dim myParam As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", sm.UserInfo.Years}}
        Dim sql As String = ""
        'sql = "" & vbCrLf
        sql &= " SELECT cc.ORGNAME,cc.OCID,cc.CLSID,cc.PCS,FORMAT(cc.MODIFYDATE,'mmssdd') MSD" & vbCrLf
        sql &= " ,cc.TPLANID,cc.PLANID,cc.COMIDNO,cc.SEQNO,cc.RID,cc.APPSTAGE" & vbCrLf
        sql &= " ,cc.YEARS,cc.THOURS,cc.TNUM" & vbCrLf
        sql &= " ,format(cc.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,format(cc.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        sql &= " ,cc.CYCLTYPE,cc.CLASSCNAME,cc.CLASSCNAME2" & vbCrLf
        sql &= " ,sb.UNLOCKSTATE" & vbCrLf
        sql &= " FROM VIEW2 cc" & vbCrLf
        sql &= " LEFT JOIN CLASS_SUBINFO sb ON sb.OCID=cc.OCID " & vbCrLf
        sql &= " WHERE cc.TPLANID=@TPLANID" & vbCrLf
        sql &= " AND cc.YEARS=@YEARS" & vbCrLf
        Select Case sm.UserInfo.LID
            Case 0
                '訓練機構
                If RIDValue.Value <> "" AndAlso RIDValue.Value.Length > 1 Then
                    myParam.Add("RID", RIDValue.Value)
                    sql &= " AND cc.RID=@RID" & vbCrLf
                End If
            Case 1
                myParam.Add("DistID", sDistID)
                myParam.Add("PlanID", sm.UserInfo.PlanID)
                sql &= " AND cc.DistID=@DistID" & vbCrLf
                sql &= " AND cc.PlanID=@PlanID" & vbCrLf
                '訓練機構
                If RIDValue.Value <> "" AndAlso RIDValue.Value.Length > 1 Then
                    myParam.Add("RID", RIDValue.Value)
                    sql &= " AND cc.RID=@RID" & vbCrLf
                End If
            Case Else
                myParam.Add("DistID", sDistID)
                myParam.Add("PlanID", sm.UserInfo.PlanID)
                sql &= " AND cc.DistID=@DistID" & vbCrLf
                sql &= " AND cc.PlanID=@PlanID" & vbCrLf
                '訓練機構
                myParam.Add("RID", RIDValue.Value)
                sql &= " AND cc.RID=@RID" & vbCrLf
        End Select
        '班級名稱
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        If ClassName.Text <> "" Then
            Dim ClassName_lk As String = String.Concat("%", ClassName.Text, "%")
            myParam.Add("ClassName_lk", ClassName_lk)
            sql &= " AND cc.CLASSCNAME LIKE @ClassName_lk" & vbCrLf
        End If
        '期別
        CyclType.Text = TIMS.ClearSQM(CyclType.Text)
        If CyclType.Text <> "" Then
            CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
            myParam.Add("CyclType", CyclType.Text)
            sql &= " AND cc.CyclType=@CyclType" & vbCrLf
        End If
        's_OCID.Text = TIMS.ClearSQM(s_OCID.Text)
        '轉換數字後相等:true/false:異常
        Dim flag_can_use_OCID As Boolean = (s_OCID.Text <> "" AndAlso TIMS.IsNumeric2(s_OCID.Text))
        If flag_can_use_OCID Then
            myParam.Add("OCID", CInt(Val(s_OCID.Text)))
            sql &= " AND cc.OCID=@OCID" & vbCrLf
        End If
        '依申請階段 
        Dim v_AppStage As String = "" 'TIMS.GetListValue(AppStage)
        If tr_AppStage_TP28.Visible Then v_AppStage = TIMS.GetListValue(AppStage)
        If v_AppStage <> "" Then
            myParam.Add("APPSTAGE", v_AppStage)
            sql &= " AND cc.APPSTAGE=@APPSTAGE" & vbCrLf
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, myParam)

        SchPanel.Visible = False
        DataGrid1.Visible = False
        msg.Text = TIMS.cst_NODATAMsg1
        If dt.Rows.Count = 0 Then Exit Sub

        SchPanel.Visible = True
        DataGrid1.Visible = True
        msg.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    ''' <summary>查詢</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnQuery_Click(sender As Object, e As EventArgs) Handles BtnQuery.Click
        s_OCID.Text = TIMS.ClearSQM(s_OCID.Text)
        If s_OCID.Text <> "" Then
            Dim drCC As DataRow = TIMS.GetOCIDDate(s_OCID.Text, objconn)
            If drCC Is Nothing Then
                Common.MessageBox(Me, "輸入課程代碼有誤!!")
                Exit Sub
            End If
        End If

        Call sSearch1()
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        Dim OCIDVal As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim v_PLANID As String = TIMS.GetMyValue(sCmdArg, "PLANID")
        Dim v_COMIDNO As String = TIMS.GetMyValue(sCmdArg, "COMIDNO")
        Dim v_SEQNO As String = TIMS.GetMyValue(sCmdArg, "SEQNO")
        Select Case e.CommandName
            Case "BTNUNLOCK"
                Call UTL_UNLOCK(OCIDVal)
                Call sSearch1()

            Case "BTNREVISE"
                Call SHOW_DETAIL1()
                Call ClearDATA1()
                Call SHOW_CLASSINFO(OCIDVal, v_PLANID, v_COMIDNO, v_SEQNO)
                Call SHOW_PLANINFO(v_PLANID, v_COMIDNO, v_SEQNO)
                'Call UTL_BTNREVISE(OCIDVal)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                'Hid_PCS'Hid_OCID'Hid_MSD'BTNUNLOCK'BTNREVISE
                Dim drv As DataRowView = e.Item.DataItem
                'Dim Hid_PCS As HiddenField = e.Item.FindControl("Hid_PCS")
                'Dim Hid_OCID As HiddenField = e.Item.FindControl("Hid_OCID")
                'Dim Hid_MSD As HiddenField = e.Item.FindControl("Hid_MSD")
                Dim BTNUNLOCK As LinkButton = e.Item.FindControl("BTNUNLOCK") '解鎖
                Dim BTNREVISE As LinkButton = e.Item.FindControl("BTNREVISE") '修改

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) ' e.Item.ItemIndex + 1 + DG_Org.PageSize * DG_Org.CurrentPageIndex
                Dim s_UNLOCKSTATE As String = Convert.ToString(drv("UNLOCKSTATE"))
                'Hid_PCS.Value = Convert.ToString(drv("PCS"))
                'Hid_OCID.Value = Convert.ToString(drv("OCID"))
                'Hid_MSD.Value = Convert.ToString(drv("MSD"))
                Dim sCmdArg As String = ""
                'TIMS.SetMyValue(sCmdArg, "PCS", drv("PCS"))
                TIMS.SetMyValue(sCmdArg, "OCID", drv("OCID"))
                TIMS.SetMyValue(sCmdArg, "PLANID", drv("PLANID"))
                TIMS.SetMyValue(sCmdArg, "COMIDNO", drv("COMIDNO"))
                TIMS.SetMyValue(sCmdArg, "SEQNO", drv("SEQNO"))

                'TIMS.SetMyValue(sCmdArg, "MSD", drv("MSD"))
                BTNUNLOCK.CommandArgument = sCmdArg '解鎖
                BTNREVISE.CommandArgument = sCmdArg '修改
                Select Case sm.UserInfo.LID
                    Case 2
                        '[解鎖]
                        BTNUNLOCK.Visible = False
                        BTNUNLOCK.Enabled = False
                        '[修改]
                        BTNREVISE.Enabled = If(s_UNLOCKSTATE = "Y", True, False)
                        If Not BTNREVISE.Enabled Then
                            TIMS.Tooltip(BTNREVISE, "解鎖後，方可修改班級聯絡資訊")
                        Else
                            TIMS.Tooltip(BTNREVISE, "", True)
                        End If
                    Case Else
                        '[解鎖]
                        BTNUNLOCK.Enabled = If(s_UNLOCKSTATE = "Y", False, True)
                        If Not BTNUNLOCK.Enabled Then
                            TIMS.Tooltip(BTNUNLOCK, "已解鎖，可修改班級聯絡資訊")
                        Else
                            TIMS.Tooltip(BTNUNLOCK, "", True)
                        End If
                        '[修改]
                        BTNREVISE.Enabled = Not BTNUNLOCK.Enabled
                        If Not BTNREVISE.Enabled Then
                            TIMS.Tooltip(BTNREVISE, "解鎖後，方可修改班級聯絡資訊")
                        Else
                            TIMS.Tooltip(BTNREVISE, "", True)
                        End If
                End Select
        End Select
    End Sub

    ''' <summary>解鎖</summary>
    ''' <param name="OCIDVal"></param>
    Sub UTL_UNLOCK(ByRef OCIDVal As String)
        Dim sParms As New Hashtable From {{"OCID", OCIDVal}}
        Dim sSql As String = "SELECT 1 FROM CLASS_SUBINFO WHERE OCID=@OCID"
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms)

        If dt.Rows.Count = 0 Then
            Dim iParms As New Hashtable From {{"OCID", OCIDVal}, {"UNLOCKACCT", sm.UserInfo.UserID}, {"MODIFYACCT", sm.UserInfo.UserID}}
            Dim i_Sql As String = ""
            i_Sql &= " INSERT INTO CLASS_SUBINFO(OCID,UNLOCKACCT,UNLOCKDATE,REVISEACCT,REVISEDATE,UNLOCKSTATE,MODIFYACCT,MODIFYDATE)" & vbCrLf
            i_Sql &= " VALUES (@OCID ,@UNLOCKACCT ,GETDATE(),NULL,NULL,'Y',@MODIFYACCT ,GETDATE())" & vbCrLf
            DbAccess.ExecuteNonQuery(i_Sql, objconn, iParms)
        Else
            Dim uParms As New Hashtable From {{"UNLOCKACCT", sm.UserInfo.UserID}, {"MODIFYACCT", sm.UserInfo.UserID}, {"OCID", OCIDVal}}
            Dim u_Sql As String = ""
            u_Sql &= " UPDATE CLASS_SUBINFO SET UNLOCKACCT=@UNLOCKACCT ,UNLOCKDATE=GETDATE() ,REVISEACCT=NULL,REVISEDATE=NULL,UNLOCKSTATE='Y'" & vbCrLf
            u_Sql &= ",MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE() WHERE OCID=@OCID" & vbCrLf
            DbAccess.ExecuteNonQuery(u_Sql, objconn, uParms)
        End If
    End Sub

    ''' <summary>帶出相關資料-課程-CLASS_CLASSINFO</summary>
    ''' <param name="rq_OCID"></param>
    ''' <param name="PlanID"></param>
    ''' <param name="ComIDNO"></param>
    ''' <param name="SeqNO"></param>
    Sub SHOW_CLASSINFO(ByVal rq_OCID As String, ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNO As String)
        'rq_OCID = TIMS.ClearSQM(rq_OCID)
        Hid_OCID.Value = ""
        If rq_OCID = "" OrElse PlanID = "" OrElse ComIDNO = "" OrElse SeqNO = "" Then Return

        Dim sParms As Hashtable = New Hashtable From {{"OCID", rq_OCID}, {"PlanID", PlanID}, {"ComIDNO", ComIDNO}, {"SeqNO", SeqNO}}
        Dim sql As String = ""
        sql = "SELECT * FROM dbo.CLASS_CLASSINFO WHERE OCID=@OCID AND PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, sParms)
        If dr Is Nothing Then Return
        Dim dr2 As DataRow = TIMS.Get_ORGINFOdr2(ComIDNO, objconn)
        If dr2 Is Nothing Then Return

        TB_OrgName.Text = Convert.ToString(dr2("ORGNAME"))

        Hid_OCID.Value = Convert.ToString(dr("OCID"))
        Hid_PlanID.Value = Convert.ToString(dr("PlanID"))
        Hid_ComIDNO.Value = Convert.ToString(dr("ComIDNO"))
        Hid_SeqNo.Value = Convert.ToString(dr("SeqNo"))

        TB_OCID.Text = Convert.ToString(dr("OCID"))
        TB_ClassCName.Text = Convert.ToString(dr("CLASSCNAME"))
        TB_CYCLTYPE.Text = TIMS.FmtCyclType(dr("CyclType"))
    End Sub

    ''' <summary>訓練計畫資料顯示--PLAN_PLANINFO</summary>
    ''' <param name="PlanID"></param>
    ''' <param name="ComIDNO"></param>
    ''' <param name="SeqNO"></param>
    Sub SHOW_PLANINFO(ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNO As String)
        'rq_OCID = TIMS.ClearSQM(rq_OCID)
        If PlanID = "" OrElse ComIDNO = "" OrElse SeqNO = "" Then Return

        Dim sParms As New Hashtable From {{"PlanID", PlanID}, {"ComIDNO", ComIDNO}, {"SeqNO", SeqNO}}
        Dim sql As String = ""
        sql = " SELECT * FROM dbo.PLAN_PLANINFO WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, sParms)
        If dr Is Nothing Then Return
        TB_ContactName.Text = dr("ContactName").ToString
        'TB_ContactPhone.Text = dr("ContactPhone").ToString
        '2023/2024/ContactPhone
        'ContactPhone.Text = Convert.ToString(dr("ContactPhone")) '電話
        Dim hCtPhone As New Hashtable
        TIMS.CHK_ContactPhoneFMT(Convert.ToString(dr("ContactPhone")), hCtPhone)
        ContactPhone_1.Text = hCtPhone("ContactPhone_1")
        ContactPhone_2.Text = hCtPhone("ContactPhone_2")
        ContactPhone_3.Text = hCtPhone("ContactPhone_3")
        Dim hCtMobile As New Hashtable
        TIMS.CHK_ContactMobileFMT(Convert.ToString(dr("ContactMobile")), hCtMobile)
        ContactMobile_1.Text = hCtMobile("ContactMobile_1")
        ContactMobile_2.Text = hCtMobile("ContactMobile_2")
        ContactPhone_1.Text = TIMS.ClearSQM(ContactPhone_1.Text)
        ContactPhone_2.Text = TIMS.ClearSQM(ContactPhone_2.Text)
        ContactPhone_3.Text = TIMS.ClearSQM(ContactPhone_3.Text)
        ContactMobile_1.Text = TIMS.ClearSQM(ContactMobile_1.Text)
        ContactMobile_2.Text = TIMS.ClearSQM(ContactMobile_2.Text)
        ' trContactPhone_2024_N1,' ContactPhone_1,' ContactPhone_2,' ContactPhone_3,' ContactMobile_1,' ContactMobile_2,' trContactPhone_2024_N2,' lab_ContactPhone_m1,
        '(【辦公室電話】、【行動電話】至少須擇一填寫)
        'lab_ContactMobile_m2
        '(【辦公室電話】、【行動電話】至少須擇一填寫)
        TB_ContactEmail.Text = dr("ContactEmail").ToString
        TB_ContactFax.Text = dr("ContactFax").ToString
    End Sub

    Sub ClearDATA1()
        TB_OrgName.Text = "" 'Convert.ToString(dr2("ORGNAME"))

        Hid_OCID.Value = "" 'Convert.ToString(dr("OCID"))
        Hid_PlanID.Value = "" 'Convert.ToString(dr("PlanID"))
        Hid_ComIDNO.Value = "" 'Convert.ToString(dr("ComIDNO"))
        Hid_SeqNo.Value = "" 'Convert.ToString(dr("SeqNo"))

        TB_OCID.Text = "" 'Convert.ToString(dr("OCID"))
        TB_ClassCName.Text = "" 'Convert.ToString(dr("CLASSCNAME"))
        TB_CYCLTYPE.Text = "" 'TIMS.FmtCyclType(dr("CyclType"))

        TB_ContactName.Text = "" 'dr("ContactName").ToString
        'TB_ContactPhone.Text = "" 'dr("ContactPhone").ToString
        ContactPhone_1.Text = ""
        ContactPhone_2.Text = ""
        ContactPhone_3.Text = ""
        ContactMobile_1.Text = ""
        ContactMobile_2.Text = ""

        TB_ContactEmail.Text = "" 'dr("ContactEmail").ToString
        TB_ContactFax.Text = "" 'dr("ContactFax").ToString
    End Sub

    Sub SAVE_REVISE(ByVal OCIDVal As String)
        OCIDVal = TIMS.ClearSQM(OCIDVal)

        Dim sParms As New Hashtable From {{"OCID", OCIDVal}}
        Dim sSql As String = "SELECT 1 FROM CLASS_SUBINFO WHERE OCID=@OCID"
        Dim dr As DataRow = DbAccess.GetOneRow(sSql, objconn, sParms)
        If dr Is Nothing Then Return

        Dim uParms As New Hashtable From {{"REVISEACCT", sm.UserInfo.UserID}, {"MODIFYACCT", sm.UserInfo.UserID}, {"OCID", OCIDVal}}
        Dim u_Sql As String = ""
        u_Sql &= " UPDATE CLASS_SUBINFO SET REVISEACCT=@REVISEACCT,REVISEDATE=GETDATE(),UNLOCKSTATE=NULL" & vbCrLf
        u_Sql &= ",MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE() WHERE OCID=@OCID" & vbCrLf
        DbAccess.ExecuteNonQuery(u_Sql, objconn, uParms)
    End Sub

    Sub SAVE_PLANINFO()
        Hid_OCID.Value = TIMS.ClearSQM(Hid_OCID.Value)
        Hid_PlanID.Value = TIMS.ClearSQM(Hid_PlanID.Value)
        Hid_ComIDNO.Value = TIMS.ClearSQM(Hid_ComIDNO.Value)
        Hid_SeqNo.Value = TIMS.ClearSQM(Hid_SeqNo.Value)

        If Hid_OCID.Value = "" OrElse Hid_PlanID.Value = "" OrElse Hid_ComIDNO.Value = "" OrElse Hid_SeqNo.Value = "" Then Return

        'Dim sParms As New Hashtable
        'sParms.Add("OCID", Hid_OCID.Value)
        'sParms.Add("PlanID", Hid_PlanID.Value)
        'sParms.Add("ComIDNO", Hid_ComIDNO.Value)
        'sParms.Add("SeqNO", Hid_SeqNo.Value)
        'Dim sql As String = ""
        'sql = "SELECT * FROM dbo.CLASS_CLASSINFO WHERE OCID=@OCID AND PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
        'Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, sParms)
        'If dr Is Nothing Then Return

        Dim sParms2 As New Hashtable From {{"PlanID", Hid_PlanID.Value}, {"ComIDNO", Hid_ComIDNO.Value}, {"SeqNO", Hid_SeqNo.Value}}
        Dim sql2 As String = ""
        sql2 = " SELECT * FROM dbo.PLAN_PLANINFO WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
        Dim dr2 As DataRow = DbAccess.GetOneRow(sql2, objconn, sParms2)
        If dr2 Is Nothing Then Return

        '2023/2024/ContactPhone
        'ContactPhone.Text = TIMS.ClearSQM(ContactPhone.Text)
        ContactPhone_1.Text = TIMS.ClearSQM(ContactPhone_1.Text)
        ContactPhone_2.Text = TIMS.ClearSQM(ContactPhone_2.Text)
        ContactPhone_3.Text = TIMS.ClearSQM(ContactPhone_3.Text)
        ContactMobile_1.Text = TIMS.ClearSQM(ContactMobile_1.Text)
        ContactMobile_2.Text = TIMS.ClearSQM(ContactMobile_2.Text)
        'Dim s_ContactPhone As String = If(fg_phone_2024, TIMS.ChangContactPhone(ContactPhone_1.Text, ContactPhone_2.Text, ContactPhone_3.Text), ContactPhone.Text)
        Dim s_ContactPhone As String = TIMS.ChangContactPhone(ContactPhone_1.Text, ContactPhone_2.Text, ContactPhone_3.Text)
        'dr("ContactPhone") = If(s_ContactPhone <> "", s_ContactPhone, Convert.DBNull)
        Dim s_ContactMobile As String = TIMS.ChangContactMobile(ContactMobile_1.Text, ContactMobile_2.Text)
        'dr("ContactMobile") = If(s_ContactMobile <> "", s_ContactMobile, Convert.DBNull)

        TB_ContactName.Text = TIMS.ClearSQM(TB_ContactName.Text)
        'TB_ContactPhone.Text = TIMS.ClearSQM(TB_ContactPhone.Text)
        TB_ContactEmail.Text = TIMS.ClearSQM(TB_ContactEmail.Text)
        TB_ContactFax.Text = TIMS.ClearSQM(TB_ContactFax.Text)
        TB_ContactEmail.Text = TIMS.ChangeEmail(TB_ContactEmail.Text)

        Dim uParms As New Hashtable From {
            {"ContactName", TB_ContactName.Text},
            {"ContactPhone", If(s_ContactPhone <> "", s_ContactPhone, Convert.DBNull)},
            {"ContactMobile", If(s_ContactMobile <> "", s_ContactMobile, Convert.DBNull)},
            {"ContactEmail", TB_ContactEmail.Text},
            {"ContactFax", TB_ContactFax.Text},
            {"MODIFYACCT", sm.UserInfo.UserID},
            {"PlanID", Hid_PlanID.Value},
            {"ComIDNO", Hid_ComIDNO.Value},
            {"SeqNO", Hid_SeqNo.Value}
        }
        Dim u_Sql As String = ""
        'u_Sql = "" & vbCrLf
        u_Sql &= " UPDATE PLAN_PLANINFO" & vbCrLf
        u_Sql &= " SET CONTACTNAME=@ContactName" & vbCrLf
        u_Sql &= " ,ContactPhone=@ContactPhone" & vbCrLf
        u_Sql &= " ,ContactMobile=@ContactMobile" & vbCrLf
        u_Sql &= " ,CONTACTEMAIL=@ContactEmail" & vbCrLf
        u_Sql &= " ,CONTACTFAX=@ContactFax" & vbCrLf
        u_Sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        u_Sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        u_Sql &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO=@SeqNO" & vbCrLf
        DbAccess.ExecuteNonQuery(u_Sql, objconn, uParms)
    End Sub

    Sub SHOW_DETAIL1()
        EdtPanel1.Visible = True
        SchPanel.Visible = False
        SchPanel2.Visible = SchPanel.Visible

        TB_OrgName.Enabled = False
        TB_OCID.Enabled = False
        TB_ClassCName.Enabled = False
        TB_CYCLTYPE.Enabled = False
    End Sub

    Sub SHOW_SEARCH1()
        EdtPanel1.Visible = False
        SchPanel.Visible = True
        SchPanel2.Visible = SchPanel.Visible
        'If Not Me.ViewState("ClassSearchStr") Is Nothing Then Session("ClassSearchStr") = Me.ViewState("ClassSearchStr")
        ''Response.Redirect("TC_01_004.aspx?ID=" & Request("ID") & "")
        'Dim url1 As String = "TC_01_004.aspx?ID=" & Request("ID") & ""
        'Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        Call ClearDATA1()
        Call SHOW_SEARCH1()
    End Sub

    Sub CheckDataPhone(ByRef ErrMsg As String)
        '【辦公室電話】、【行動電話】至少須擇一填寫
        '2023/2024/ContactPhone
        ContactPhone_1.Text = TIMS.ClearSQM(ContactPhone_1.Text)
        ContactPhone_2.Text = TIMS.ClearSQM(ContactPhone_2.Text)
        ContactPhone_3.Text = TIMS.ClearSQM(ContactPhone_3.Text)
        ContactMobile_1.Text = TIMS.ClearSQM(ContactMobile_1.Text)
        ContactMobile_2.Text = TIMS.ClearSQM(ContactMobile_2.Text)
        Dim s_ContactPhone As String = TIMS.ChangContactPhone(ContactPhone_1.Text, ContactPhone_2.Text, ContactPhone_3.Text)
        Dim s_ContactMobile As String = TIMS.ChangContactMobile(ContactMobile_1.Text, ContactMobile_2.Text)
        If s_ContactPhone = "" AndAlso s_ContactMobile = "" Then ErrMsg &= "請輸入 班別資料-【辦公室電話】、【行動電話】至少須擇一填寫" & vbCrLf

        If s_ContactPhone <> "" AndAlso ContactPhone_1.Text <> "" AndAlso Not TIMS.IsNumberStr(ContactPhone_1.Text) Then
            ErrMsg &= "班別資料-【辦公室電話】區碼，僅能為數字" & vbCrLf
        ElseIf s_ContactPhone <> "" AndAlso ContactPhone_1.Text <> "" AndAlso Not ContactPhone_1.Text.StartsWith("0") Then
            ErrMsg &= "班別資料-【辦公室電話】區碼，第1碼應該為0" & vbCrLf
        ElseIf s_ContactPhone <> "" AndAlso ContactPhone_1.Text <> "" AndAlso ContactPhone_1.Text.Length < 2 Then
            ErrMsg &= "班別資料-【辦公室電話】區碼，長度須大於1" & vbCrLf
        ElseIf s_ContactPhone <> "" AndAlso ContactPhone_2.Text <> "" AndAlso Not TIMS.IsNumberStr(ContactPhone_2.Text) Then
            ErrMsg &= "班別資料-【辦公室電話】電話(8碼)，僅能為數字" & vbCrLf
        ElseIf s_ContactPhone <> "" AndAlso ContactPhone_3.Text <> "" AndAlso Not TIMS.IsNumberStr(ContactPhone_3.Text) Then
            ErrMsg &= "班別資料-【辦公室電話】分機，僅能為數字" & vbCrLf
        ElseIf s_ContactPhone <> "" AndAlso (ContactPhone_1.Text = "" OrElse ContactPhone_2.Text = "") Then
            ErrMsg &= "班別資料-【辦公室電話】不為空，請填寫完整(區碼與電話)為必填" & vbCrLf
        End If

        If s_ContactMobile <> "" AndAlso ContactMobile_1.Text <> "" AndAlso Not TIMS.IsNumberStr(ContactMobile_1.Text) Then
            ErrMsg &= "班別資料-【行動電話】手機前4碼，僅能為數字" & vbCrLf
        ElseIf s_ContactMobile <> "" AndAlso ContactMobile_1.Text <> "" AndAlso Not ContactMobile_1.Text.StartsWith("0") Then
            ErrMsg &= "班別資料-【行動電話】手機前4碼，第1碼應該為0" & vbCrLf
        ElseIf s_ContactMobile <> "" AndAlso ContactMobile_2.Text <> "" AndAlso Not TIMS.IsNumberStr(ContactMobile_2.Text) Then
            ErrMsg &= "班別資料-【行動電話】手機後6碼，僅能為數字" & vbCrLf
        ElseIf s_ContactMobile <> "" AndAlso (ContactMobile_1.Text = "" OrElse ContactMobile_2.Text = "") Then
            ErrMsg &= "班別資料-【行動電話】不為空，請填寫完整(前4碼與後6碼)為必填" & vbCrLf
        End If
    End Sub

    Protected Sub BtnSAVEDATA1_Click(sender As Object, e As EventArgs) Handles BtnSAVEDATA1.Click
        Dim sErrMsg As String = ""
        CheckDataPhone(sErrMsg)
        If sErrMsg <> "" Then
            '有錯誤訊息
            'sm.LastErrorMessage = sErrMsg
            Common.MessageBox(Me, sErrMsg)
            Return 'Exit Sub 'Return False '不可儲存
        End If

        Call SAVE_PLANINFO()  '儲存(PLAN_PLANINFO)
        Call SAVE_REVISE(Hid_OCID.Value)
        Common.MessageBox(Me, "儲存完畢")
        Call ClearDATA1()
        Call SHOW_SEARCH1()
        Call sSearch1()
    End Sub
End Class