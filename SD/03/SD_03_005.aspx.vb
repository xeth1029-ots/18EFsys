Partial Class SD_03_005
    Inherits AuthBasePage

    Dim vMsg1 As String = ""
    Const cst_msg1 As String = "委訓單位，不允許學員刪除作業!!"

    Const cst_dr擅打錯誤 As String = "1"
    Const cst_dr資格不符 As String = "2"
    Const cst_dr不符參訓資格 As String = "4"
    Const cst_dr其他 As String = "3"

    'https://jira.turbotech.com.tw/browse/TIMSC-187
    '署使用者，刪除學員，則不受學員已有其他相關資料的限制
    'Dim gflagLID0CanDel As Boolean = False

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在---------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線
        '檢查Session是否存在---------------------------End

        If Not xChk_LID2NGDel() Then Exit Sub '該權限不可執行刪除功能。
        If Not IsPostBack Then Call sCreate1()

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True, "Button1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

    End Sub

    '該權限不可執行刪除功能。
    Function xChk_LID2NGDel() As Boolean
        If Not TIMS.Chk_LID2NGDel(Me, objconn) Then
            msg.Text = ""
            DataGridTable.Visible = False
            Button2.Enabled = False
            Button1.Enabled = False
            Common.MessageBox(Me, cst_msg1)
            Return False
        End If
        'https://jira.turbotech.com.tw/browse/TIMSC-187
        'gflagLID0CanDel = TIMS.Chk_LID0CanDel(Me)
        Return True
    End Function

    Sub sCreate1()
        msg.Text = ""
        DataGridTable.Visible = False
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        Button1.Attributes("onclick") = "javascript:return search()" '查詢
        'Button1.Attributes("onclick") = "return search();"
        Button2.Attributes("onclick") = "return CheckData();"
        'Button2.Enabled = False
        'If au.blnCanDel Then Button2.Enabled = True '刪除學員
        'Button1.Enabled = False
        'If au.blnCanSech Then Button1.Enabled = True '查詢
    End Sub

    '查詢
    Sub sSearch1()

        Dim parms As New Hashtable
        parms.Add("OCID", Val(OCIDValue1.Value))

        Dim sql As String = ""
        sql &= " SELECT a.SOCID" & vbCrLf
        sql &= " ,a.StudentID" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(a.StudentID) STUDID2" & vbCrLf
        sql &= " ,a.StudStatus" & vbCrLf
        sql &= " ,dbo.DECODE12(a.StudStatus,1,'在訓',2,'離訓',3,'退訓',4,'續訓',5,'結訓','在訓') StudStatus2" & vbCrLf
        sql &= " ,b.Sex" & vbCrLf
        sql &= " ,dbo.DECODE6(b.Sex,'M','男','F','女',b.Sex) Sex2" & vbCrLf
        sql &= " ,b.NAME" & vbCrLf
        sql &= " ,b.IDNO ,dbo.FN_GET_MASK1(b.IDNO) IDNO_MK" & vbCrLf
        sql &= " ,format(b.BIRTHDAY,'yyyy/MM/dd') BIRTHDAY,'**/**/**' BIRTHDAY_MK" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS a" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO b ON a.SID = b.SID" & vbCrLf
        sql &= " WHERE a.OCID = @OCID" & vbCrLf
        sql &= " ORDER BY a.StudentID" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        msg.Text = "查無資料"
        DataGridTable.Visible = False

        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Visible = True
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                'e.Item.CssClass = "SD_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                'If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SD_TD2"
                Dim drv As DataRowView = e.Item.DataItem
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim DelResaon As DropDownList = e.Item.FindControl("DelResaon")
                Dim DelReasonOther As TextBox = e.Item.FindControl("DelReasonOther")
                Dim Label1 As Label = e.Item.FindControl("Label1")
                'https://jira.turbotech.com.tw/browse/TIMSC-187
                DelResaon = TIMS.Get_DelResaonDDL(DelResaon)
                Checkbox1.Value = Convert.ToString(drv("SOCID"))

                Dim sObj2 As String = ""
                sObj2 = ""
                sObj2 &= "'" & DelResaon.ClientID & "'"
                sObj2 &= ",'" & DelReasonOther.ClientID & "'"
                sObj2 &= ",'" & Label1.ClientID & "'"
                DelResaon.Attributes("onchange") = "SetChgDelRes1(" & sObj2 & ")"
                Select Case DelResaon.SelectedValue
                    Case cst_dr其他
                        DelReasonOther.Style("display") = ""
                        Label1.Style("display") = ""
                    Case Else
                        DelReasonOther.Style("display") = "none"
                        Label1.Style("display") = "none"
                End Select

                'https://jira.turbotech.com.tw/browse/TIMSC-187
                '檢查是否有相關紀錄-生活津貼!
                Dim dt1 As DataTable = GetSubData1(Convert.ToString(drv("SOCID")))
                If dt1.Rows.Count > 0 Then
                    'Checkbox1.Disabled = True
                    'DelResaon.Enabled = False
                    'DelReasonOther.Enabled = False
                    vMsg1 = "此學員已有"
                    For Each dr1 As DataRow In dt1.Rows
                        vMsg1 &= " " & dr1("Reason").ToString & " "
                    Next
                    'vMsg1 &= "不能刪除"
                    'TIMS.Tooltip(e.Item, vMsg1)
                    vMsg1 &= ",請謹慎刪除"
                    TIMS.Tooltip(Checkbox1, vMsg1)
                End If

                'https://jira.turbotech.com.tw/browse/TIMSC-187
                '署使用者，刪除學員，則不受學員已有其他相關資料的限制
                'If Not gflagLID0CanDel Then
                '檢查是否有相關紀錄
                Dim dt2 As DataTable = GetStudData2(Convert.ToString(drv("SOCID")))
                If dt2.Rows.Count > 0 Then
                    'Checkbox1.Disabled = True
                    'DelResaon.Enabled = False
                    'DelReasonOther.Enabled = False
                    vMsg1 = "此學員已有"
                    For Each dr2 As DataRow In dt2.Rows
                        vMsg1 &= " " & dr2("Reason").ToString & " "
                    Next
                    'vMsg1 &= "不能刪除"
                    'TIMS.Tooltip(e.Item, vMsg1)
                    vMsg1 &= ",請謹慎刪除"
                    TIMS.Tooltip(Checkbox1, vMsg1)
                End If
                'End If
                e.Item.Style("CURSOR") = "hand"
        End Select
    End Sub

    '生活津貼 學員相關table資料
    Function GetSubData1(ByVal dSOCID As String) As DataTable
        'Dim rst As Boolean = False 'true:有資料 false:無資料
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT '津貼資料' Reason FROM Stud_SubsidyResult WHERE SOCID = '" & dSOCID & "' "
        sql &= " UNION SELECT '職訓生活津貼資料' Reason FROM Sub_SubSidyApply WHERE SOCID = '" & dSOCID & "' "
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        Return dt
    End Function

    'tims學員相關table資料
    Function GetStudData2(ByVal dSOCID As String) As DataTable
        dSOCID = TIMS.ClearSQM(dSOCID)
        Dim sql As String = ""
        sql &= " SELECT '技能檢定資料' Reason FROM Stud_TechExam WHERE SOCID = '" & dSOCID & "' "
        sql &= " UNION SELECT '結訓成績資料' Reason FROM Stud_TrainingResults WHERE SOCID = '" & dSOCID & "' AND Results > 0 " '分數要大於零
        sql &= " UNION SELECT '操行成績資料' Reason FROM Stud_Conduct WHERE SOCID = '" & dSOCID & "' "
        sql &= " UNION SELECT '轉班資料' Reason FROM Stud_TranClassRecord WHERE SOCID = '" & dSOCID & "' "
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            sql &= " UNION SELECT '出缺勤資料' Reason FROM Stud_Turnout2 WHERE SOCID='" & dSOCID & "'"
        Else
            sql &= " UNION SELECT '出缺勤資料' Reason FROM Stud_Turnout WHERE SOCID='" & dSOCID & "'"
        End If
        sql &= " UNION SELECT '獎懲資料' Reason FROM Stud_Sanction WHERE SOCID='" & dSOCID & "'"
        sql &= " UNION SELECT '學員資料卡' Reason FROM Stud_ResultStudData WHERE SOCID='" & dSOCID & "'"
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then sql &= " UNION SELECT '得到學分' Reason FROM Class_StudentsOfClass WHERE SOCID = '" & dSOCID & "' AND CreditPoints = 1 "
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        Return dt
    End Function

    '送出前檢核 'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True
        Errmsg = ""
        For Each item As DataGridItem In DataGrid1.Items
            Dim Checkbox1 As HtmlInputCheckBox = item.FindControl("Checkbox1")
            Dim DelResaon As DropDownList = item.FindControl("DelResaon")
            Dim DelReasonOther As TextBox = item.FindControl("DelReasonOther")

            If Checkbox1.Checked AndAlso Checkbox1.Value <> "" Then
                Select Case DelResaon.SelectedValue
                    Case cst_dr擅打錯誤, cst_dr資格不符, cst_dr不符參訓資格
                    Case cst_dr其他 '3
                        If DelReasonOther.Text = "" Then
                            Errmsg &= "請輸入其他原因!!" & vbCrLf
                            Return False
                        Else
                            If DelReasonOther.Text.Length > 50 Then
                                Errmsg &= "刪除原因不得輸入超過50個中文字元!!" & vbCrLf
                                Return False
                            End If
                        End If
                    Case Else
                        Errmsg &= "請選擇原因!!" & vbCrLf
                        Return False
                End Select
            End If
        Next

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    'update 甄試結果試算檔Stud_SelResult
    Public Shared Sub UPD_SELRESULT(ByVal ss3 As String, tConn As SqlConnection, trans As SqlTransaction)
        Dim IDNO1 As String = TIMS.GetMyValue(ss3, "IDNO1") 'dr2("SETID")
        Dim OCID1 As String = TIMS.GetMyValue(ss3, "OCID1") 'dr2("SETID")

        Dim dr2 As DataRow = Nothing
        Dim flagCanStudSelResultUpdate As Boolean = False
        Dim sql As String = ""
        sql &= " SELECT r.SETID ,CONVERT(varchar, r.ENTERDATE, 111) ETEnterDate ,r.SERNUM ,r.OCID" & vbCrLf
        sql &= " FROM STUD_SELRESULT r" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE b ON b.setid = r.setid AND b.enterdate = r.enterdate AND b.sernum = r.sernum AND b.ocid1 = r.ocid" & vbCrLf
        sql &= " JOIN STUD_ENTERTEMP a ON a.setid = b.setid" & vbCrLf
        sql &= " WHERE a.IDNO=@IDNO AND b.OCID1=@OCID1" & vbCrLf
        Dim sCmd2 As New SqlCommand(sql, tConn, trans)
        Dim dt2 As New DataTable
        With sCmd2
            .Parameters.Clear()
            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = IDNO1 'xIDNO
            .Parameters.Add("OCID1", SqlDbType.Int).Value = OCID1 'xOCID
            dt2.Load(.ExecuteReader())
        End With
        If dt2.Rows.Count > 0 Then
            dr2 = dt2.Rows(0)
            flagCanStudSelResultUpdate = True
        End If

#Region "(No Use)"

        'update 甄試結果試算檔Stud_SelResult
        'sql = "select * from Stud_SelResult WHERE SETID = '" & dr("SETID") & "' and  EnterDate = to_date('" & dr("ETEnterDate") & "','MM/DD/YYYY') and SerNum = '" & dr("SerNum") & "' and OCID = '" & dr("OCID") & "' " '★
        'Dim sql As String = ""
        'Dim sqlStr As String = ""

#End Region

        If flagCanStudSelResultUpdate Then
            Dim sqlStr As String = ""
            sqlStr &= " UPDATE Stud_SelResult"
            sqlStr &= " SET AppliedStatus='N'"
            sqlStr &= " WHERE SETID = @SETID AND EnterDate = @EnterDate AND SerNum = @SerNum  AND OCID = @OCID"
            Dim uCmd As New SqlCommand(sqlStr, tConn, trans)
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("SETID", SqlDbType.Int).Value = dr2("SETID")
                .Parameters.Add("EnterDate", SqlDbType.DateTime).Value = CDate(dr2("ETEnterDate"))
                .Parameters.Add("SerNum", SqlDbType.Int).Value = dr2("SerNum")
                .Parameters.Add("OCID", SqlDbType.Int).Value = dr2("OCID")
                '.ExecuteNonQuery()  'edit，by:20181017
                DbAccess.ExecuteNonQuery(uCmd.CommandText, trans, uCmd.Parameters)  'edit，by:20181017
            End With
        End If

#Region "(No Use)"

        'Dim SETID As String = TIMS.GetMyValue(ss3, "SETID") 'dr2("SETID")
        'Dim EnterDate As String = TIMS.GetMyValue(ss3, "EnterDate") 'CDate(dr2("ETEnterDate"))
        'Dim SerNum As String = TIMS.GetMyValue(ss3, "SerNum") 'dr2("SerNum")
        'Dim OCID As String = TIMS.GetMyValue(ss3, "OCID") 'dr2("OCID")

#End Region
    End Sub

    'DELETE FUNC1
    Sub sUtl_DeleteData1()
        '2006/03/28 add conn by matt
        Dim sql As String = ""
        Dim tConn As SqlConnection = DbAccess.GetConnection()
        Dim trans As SqlTransaction = DbAccess.BeginTrans(tConn)

        Try
            '2006/03/28 add conn by matt
            'trans = DbAccess.BeginTrans(tConn)
            For Each item As DataGridItem In DataGrid1.Items
                Dim Checkbox1 As HtmlInputCheckBox = item.FindControl("Checkbox1")
                Dim DelResaon As DropDownList = item.FindControl("DelResaon")
                Dim DelReasonOther As TextBox = item.FindControl("DelReasonOther")

                If Checkbox1.Checked AndAlso Checkbox1.Value <> "" Then
                    Dim rqSOCID As String = TIMS.ClearSQM(Checkbox1.Value)
                    If rqSOCID = "" Then Exit For

                    sql = ""
                    sql &= " SELECT f.OrgID "
                    sql &= " ,f.OrgName "
                    sql &= " ,d.PlanName "
                    sql &= " ,a.StudentID "
                    sql &= " ,dbo.FN_CSTUDID2(a.StudentID) STUDID2" & vbCrLf
                    sql &= " ,e.NAME "
                    sql &= " ,e.IDNO ,dbo.FN_GET_MASK1(e.IDNO) IDNO_MK" & vbCrLf
                    sql &= " ,a.StudStatus "
                    sql &= " ,dbo.DECODE12(a.StudStatus,1,'在訓',2,'離訓',3,'退訓',4,'續訓',5,'結訓','在訓') StudStatus2" & vbCrLf
                    sql &= " ,b.PlanID,b.ComIDNO,b.SeqNo "
                    sql &= " ,b.ClassCName "
                    sql &= " ,b.RID,a.OCID "
                    sql &= " ,a.SOCID,e.SID "
                    sql &= " FROM CLASS_STUDENTSOFCLASS a "
                    sql &= " JOIN CLASS_CLASSINFO b ON a.OCID = b.OCID "
                    sql &= " JOIN ID_PLAN c ON b.PlanID = c.PlanID "
                    sql &= " JOIN KEY_PLAN d ON c.TPlanID = d.TPlanID "
                    sql &= " JOIN STUD_STUDENTINFO e ON a.SID = e.SID "
                    sql &= " JOIN ORG_ORGINFO f ON b.ComIDNO = f.ComIDNO "
                    sql &= " WHERE a.SOCID = '" & rqSOCID & "' "

                    Dim dr As DataRow = DbAccess.GetOneRow(sql, trans)
                    If dr Is Nothing Then Exit For

                    Dim xIDNO As String = Convert.ToString(dr("IDNO"))
                    Dim xOCID As String = Convert.ToString(dr("OCID"))

                    'IIf(DelResaon.SelectedIndex = DelResaon.Items.Count - 1, ":" & DelReasonOther.Text, "")
                    Dim sDelResaonO As String = TIMS.ClearSQM(DelResaon.SelectedItem.Text)
                    Select Case DelResaon.SelectedValue
                        Case cst_dr其他
                            sDelResaonO &= ":" & TIMS.ClearSQM(DelReasonOther.Text)
                    End Select

                    Dim sbDelNote As New StringBuilder
                    sbDelNote.AppendFormat("刪除[{0}]", dr("PlanName"))
                    sbDelNote.AppendFormat("-[{0}]", dr("OrgName"))
                    sbDelNote.AppendFormat("-[{0}]", dr("ClassCName"))
                    sbDelNote.AppendFormat("-[({0}){1}]", dr("StudentID"), dr("Name"))
                    sbDelNote.AppendFormat("-[{0}]", dr("StudStatus2"))
                    sbDelNote.AppendFormat("-[{0}]", sDelResaonO)
                    'sDelNote = "刪除[" & dr("PlanName") & "]-[" & dr("OrgName") & "]-[" & dr("ClassCName") & "]-[(" & dr("StudentID") & ")" & dr("Name") & "]-[" & dr("StudStatus2") & "]-[" & sDelResaonO & "]"

                    Dim MRqID As String = TIMS.ClearSQM(Request("ID"))
                    Dim htSS As New Hashtable 'htSS Hashtable() '
                    htSS.Add("UserID", Convert.ToString(sm.UserInfo.UserID))
                    htSS.Add("iFunID", Val(MRqID))
                    htSS.Add("DistID", Convert.ToString(sm.UserInfo.DistID))
                    htSS.Add("DelNote", sbDelNote.ToString())

                    htSS.Add("iOrgID", dr("OrgID"))
                    htSS.Add("RID", Convert.ToString(dr("RID")))
                    htSS.Add("iPlanID", Convert.ToString(dr("PlanID")))
                    htSS.Add("ComIDNO", Convert.ToString(dr("ComIDNO")))
                    htSS.Add("iSeqNO", Convert.ToString(dr("SeqNO")))

                    htSS.Add("iOCID", Convert.ToString(dr("OCID")))
                    htSS.Add("iSOCID", Convert.ToString(rqSOCID))
                    'DelResaon: 4:不符參訓資格
                    htSS.Add("DelResaon", DelResaon.SelectedValue)
                    htSS.Add("DelReasonOther", DelReasonOther.Text)
                    htSS.Add("SID", Convert.ToString(dr("SID")))
                    htSS.Add("IDNO", Convert.ToString(dr("IDNO")))
                    htSS.Add("NAME", Convert.ToString(dr("NAME")))

                    'INSERT DELETE LOG 刪除LOG記錄 / (SYS_DELLOG)
                    Call TIMS.InsertDelLog(Me, htSS, objconn)

                    '刪除學員資料 (CLASS_STUDENTSOFCLASSDELDATA) CLASS_STUDENTSOFCLASS
                    Dim iRstDS As Integer = TIMS.sUtl_DelSTUDENTSOFCLASS(Me, rqSOCID, tConn, trans)

                    Dim ss3 As String = ""
                    TIMS.SetMyValue(ss3, "IDNO1", xIDNO)
                    TIMS.SetMyValue(ss3, "OCID1", xOCID)
                    Call UPD_SELRESULT(ss3, tConn, trans)

                    Dim sqlStr As String = ""
                    Dim da As SqlDataAdapter = Nothing
                    Dim dt As DataTable = Nothing
                    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        '變更Modify人時
                        'Dim sqlStr As String = ""
                        sqlStr = " UPDATE STUD_SERVICEPLACE SET MODIFYACCT = @ModifyAcct ,MODIFYDATE = GETDATE() WHERE SOCID = @SOCID "
                        Dim sqlAdp As New SqlDataAdapter
                        With sqlAdp
                            .UpdateCommand = New SqlCommand(sqlStr, tConn, trans)
                            .UpdateCommand.Parameters.Clear()
                            .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                            .UpdateCommand.Parameters.Add("SOCID", SqlDbType.Int).Value = rqSOCID
                            '.UpdateCommand.ExecuteNonQuery()  'edit，by:20181017
                            DbAccess.ExecuteNonQuery(sqlStr, trans, sqlAdp.UpdateCommand.Parameters)  'edit，by:20181017
                        End With

                        '搬移資料
                        sqlStr = " INSERT INTO STUD_SERVICEPLACEDELDATA SELECT * FROM STUD_SERVICEPLACE WHERE SOCID = @SOCID "
                        With sqlAdp
                            .InsertCommand = New SqlCommand(sqlStr, tConn, trans)
                            .InsertCommand.Parameters.Clear()
                            .InsertCommand.Parameters.Add("SOCID", SqlDbType.Int).Value = rqSOCID
                            '.InsertCommand.ExecuteNonQuery()  'edit，by:20181017
                            DbAccess.ExecuteNonQuery(sqlStr, trans, sqlAdp.InsertCommand.Parameters)  'edit，by:20181017
                        End With

                        '刪除Class_StudentsOfClass
                        sqlStr = " DELETE STUD_SERVICEPLACE WHERE SOCID = @SOCID "
                        With sqlAdp
                            .DeleteCommand = New SqlCommand(sqlStr, tConn, trans)
                            .DeleteCommand.Parameters.Clear()
                            .DeleteCommand.Parameters.Add("SOCID", SqlDbType.Int).Value = rqSOCID
                            '.DeleteCommand.ExecuteNonQuery()  'edit，by:20181017
                            DbAccess.ExecuteNonQuery(sqlStr, trans, sqlAdp.DeleteCommand.Parameters)  'edit，by:20181017
                        End With

                        '學員參訓背景(產學訓)
                        sql = " SELECT * FROM STUD_TRAINBG WHERE SOCID = '" & rqSOCID & "' "
                        dr = DbAccess.GetOneRow(sql, trans)

                        '學員參訓背景(產學訓) (刪除檔)
                        sql = " SELECT * FROM STUD_DELTRAINBG WHERE 1<>1 "
                        dt = DbAccess.GetDataTable(sql, da, trans)
                        If dr IsNot Nothing Then
                            Dim dr1 As DataRow = dt.NewRow
                            dt.Rows.Add(dr1)
                            dr1("SOCID") = dr("SOCID")
                            dr1("Q1") = dr("Q1")
                            dr1("Q3") = dr("Q3")
                            dr1("Q3_Other") = dr("Q3_Other")
                            dr1("Q4") = dr("Q4")
                            dr1("Q5") = dr("Q5")
                            dr1("Q61") = dr("Q61")
                            dr1("Q62") = dr("Q62")
                            dr1("Q63") = dr("Q63")
                            dr1("Q64") = dr("Q64")
                            dr1("ModifyAcct") = sm.UserInfo.UserID
                            dr1("ModifyDate") = Now
                            DbAccess.UpdateDataTable(dt, da, trans)

                            sql = " DELETE STUD_TRAINBG WHERE SOCID = '" & rqSOCID & "' "
                            DbAccess.ExecuteNonQuery(sql, trans)
                        End If

                        '學員參訓背景(產學訓) (刪除檔)
                        sql = " INSERT INTO STUD_DELTRAINBGQ2 (SOCID,Q2) SELECT SOCID ,Q2 FROM STUD_TRAINBGQ2 WHERE SOCID = '" & rqSOCID & "' "
                        DbAccess.ExecuteNonQuery(sql, trans)

                        '學員參訓背景(產學訓) 
                        sql = " DELETE STUD_TRAINBGQ2 WHERE SOCID = '" & rqSOCID & "' "
                        DbAccess.ExecuteNonQuery(sql, trans)
                    End If

                    'select socid ,count(1) cnt from Adp_DGTRNData group by socid having count(1)>1
                    'select socid ,count(1) cnt from Adp_TRNData group by socid having count(1)>1
                    'select socid ,count(1) cnt from Adp_GOVTRNData group by socid having count(1) > 1
                    '消除三合一資料
                    '學習券
                    sql = "SELECT * FROM ADP_DGTRNDATA WHERE SOCID = '" & rqSOCID & "' "
                    dt = DbAccess.GetDataTable(sql, da, trans)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("SOCID") = Convert.DBNull
                        dr("ARVL_STATE") = 0
                        dr("ARVL_DATE") = Convert.DBNull
                        dr("ARVL_UNIT_NAME") = Convert.DBNull
                        dr("ARVL_ORG_NAME") = Convert.DBNull
                        dr("ARVL_ORG_DOCNO") = Convert.DBNull
                        dr("ARVL_SDATE") = Convert.DBNull
                        dr("ARVL_EDATE") = Convert.DBNull
                        dr("ARVL_UNIT_PROMOTER") = Convert.DBNull
                        dr("ARVL_UNIT_TEL") = Convert.DBNull
                        dr("ACT_END_DATE") = Convert.DBNull
                        dr("ARVL_CLASS_NAME") = Convert.DBNull
                        dr("ARVL_CLASS_NO") = Convert.DBNull
                        dr("TransToTIMS") = "N"
                        dr("TIMSModifyDate") = Now
                        DbAccess.UpdateDataTable(dt, da, trans)
                    End If

                    '職訓券
                    sql = " SELECT * FROM ADP_TRNDATA WHERE SOCID = '" & rqSOCID & "' "
                    dt = DbAccess.GetDataTable(sql, da, trans)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("SOCID") = Convert.DBNull
                        dr("ARVL_STATE") = 0
                        dr("ARVL_DATE") = Convert.DBNull
                        dr("ARVL_SDATE") = Convert.DBNull
                        dr("ARVL_EDATE") = Convert.DBNull
                        dr("ARVL_UNIT_NAME") = Convert.DBNull
                        dr("ARVL_HOURS") = Convert.DBNull
                        dr("ARVL_CLASS_NAME") = Convert.DBNull
                        dr("ARVL_UNIT_ZIP") = Convert.DBNull
                        dr("ARVL_UNIT_ADDR") = Convert.DBNull
                        dr("ARVL_UNIT_USER") = Convert.DBNull
                        dr("ARVL_UNIT_TEL") = Convert.DBNull
                        dr("ARVL_FSH_DATE") = Convert.DBNull
                        dr("TransToTIMS") = "N"
                        dr("TIMSModifyDate") = Now
                        DbAccess.UpdateDataTable(dt, da, trans)
                    End If

                    '推介券
                    sql = " SELECT * FROM ADP_GOVTRNDATA WHERE SOCID = '" & rqSOCID & "' "
                    dt = DbAccess.GetDataTable(sql, da, trans)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("SOCID") = Convert.DBNull
                        dr("ARVL_STATE") = 0
                        dr("ARVL_DATE") = Convert.DBNull
                        dr("ARVL_EDATE") = Convert.DBNull
                        dr("ARVL_SDATE") = Convert.DBNull
                        dr("ARVL_UNIT_NAME") = Convert.DBNull
                        dr("ARVL_HOURS") = Convert.DBNull
                        dr("ARVL_CLASS_NAME") = Convert.DBNull
                        dr("ARVL_UNIT_ZIP") = Convert.DBNull
                        dr("ARVL_UNIT_ADDR") = Convert.DBNull
                        dr("ARVL_UNIT_USER") = Convert.DBNull
                        dr("ARVL_UNIT_TEL") = Convert.DBNull
                        dr("ARVL_FSH_DATE") = Convert.DBNull
                        dr("TransToTIMS") = "N"
                        dr("TIMSModifyDate") = Now
                        DbAccess.UpdateDataTable(dt, da, trans)
                    End If
                End If
            Next
            DbAccess.CommitTrans(trans)
            'TIMS.CloseDbConn(tConn)
        Catch ex As Exception
            'If Not trans Is Nothing Then DbAccess.RollbackTrans(trans)
            DbAccess.RollbackTrans(trans)
            TIMS.CloseDbConn(tConn)
            Dim strErrmsg As String = ""
            strErrmsg &= "ex.ToString:" & vbCrLf
            strErrmsg &= ex.ToString & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            Common.MessageBox(Me, TIMS.cst_NODATAMsg3)
            Exit Sub
        End Try
        TIMS.CloseDbConn(tConn)
        Common.MessageBox(Me, "刪除成功")
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not xChk_LID2NGDel() Then Exit Sub '該權限不可執行刪除功能。
        Call sSearch1()
    End Sub

    '刪除學員
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Not xChk_LID2NGDel() Then Exit Sub '該權限不可執行刪除功能。
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call sUtl_DeleteData1()
        Call sSearch1()
    End Sub

    '判斷機構是否只有一個班級
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn) '判斷機構是否只有一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Visible = False
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGridTable.Visible = False
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

End Class