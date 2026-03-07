Partial Class SD_02_001
    Inherits AuthBasePage

    'use Sys_GlobalVar
    'update STUD_ENTERTYPE
    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False
    Dim flag_File1_csv As Boolean = False

    Const cst_xc_編號 As Integer = 4
    Const cst_xc_考生姓名 As Integer = 5
    Const cst_xc_科目代號 As Integer = 6
    Const cst_xc_缺考 As Integer = 7
    Const cst_xc_成績 As Integer = 8
    Const cst_xc_口試 As Integer = 110
    Const cst_xc_Len As Integer = 111
    Const cst_xc_Len2 As Integer = 110

    'Button2 (此按鈕，無儲存動作)
    Const cst_Button2tip As String = "(系統協助總成績試算)此按鈕，無儲存動作"
    'Cells 'Columns
    Const cst_准考證號 As Integer = 0
    Const cst_姓名 As Integer = 1
    Const cst_身分證號碼 As Integer = 2
    Const cst_報名日 As Integer = 3

    Const cst_筆試成績 As Integer = 4
    Const cst_口試成績 As Integer = 5
    Const cst_總成績 As Integer = 6

    'Const cst_甄試加分 As Integer = 6 '甄試加分(加權3%) EXAMPLUS
    'Const cst_身分別 As Integer = 7
    'Const cst_總成績 As Integer = 8
    'Const cst_券別 As Integer = 9

    '匯入使用:
    Const cst_a准考證號碼 As Integer = 0
    Const cst_a編號 As Integer = 1
    Const cst_a姓名 As Integer = 2
    Const cst_a身分證號碼 As Integer = 3
    Const cst_a報名日期 As Integer = 4
    Const cst_a筆試成績 As Integer = 5
    Const cst_a口試成績 As Integer = 6
    Const cst_a總成績 As Integer = 7

    'Const cst_a加權 As Integer = 7
    'Const cst_a身分別代碼 As Integer = 8
    'Const cst_a總成績 As Integer = 9

    'Const cst_msg219 As String = "※ 姓名前標記「x-」表示民眾已註銷推介"
    'Const cst_fgb219 As String = "x-"
    'Const cst_Mgc219 As String = "民眾已註銷推介"
    'Dim blnP0 As Boolean = False '報名管道(職前計畫顯示)

    Const cst_免筆試 As String = "此班級無筆試甄試方式"
    Const cst_免口試 As String = "此班級無口試甄試方式"
    Const cst_imp_免筆試 As String = "此班級無筆試甄試方式，故系統不會匯入筆試成績資料。"
    Const cst_imp_免口試 As String = "此班級無口試甄試方式，故系統不會匯入口試成績資料。"

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        'Dim PlanID As String
        'PlanID = sm.UserInfo.PlanID

        'blnP0 = TIMS.Get_TPlanID_P0(Me, objconn)
        Trwork2013a.Visible = False '報名管道(職前計畫顯示)
        'If blnP0 Then Trwork2013a.Visible = True

        ''就服單位協助報名
        'Trwork2013a.Visible = False
        'If sm.UserInfo.Years >= 2013 AndAlso TIMS.Cst_TPlanID0237AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    If TIMS.Utl_GetConfigSet("work2013") = "Y" Then Trwork2013a.Visible = True
        'End If
        LinkButton1.Attributes("onclick") = "ShowEStudList(this);return false;"

        'Button2.Enabled = True
        'Button3.Enabled = True
        'If Not au.blnCanAdds Then
        '    Button2.Enabled = False
        '    TIMS.Tooltip(Button2, "無新增權限")
        '    Button3.Enabled = False
        '    TIMS.Tooltip(Button3, "無新增權限")
        'End If

        'Button1.Enabled = True
        'If Not au.blnCanSech Then
        '    Button1.Enabled = False
        '    TIMS.Tooltip(Button1, "無查詢權限")
        'End If

        '職訓卷
        'SELECT * FROM STUD_ENTERTEMP WHERE IDNO ='F121096128' AND SETID='1094382'
        'SELECT EnterPath FROM STUD_ENTERTYPE WHERE SETID='1094382'
        'SELECT * FORM KEY_PLAN WHERE TPLANID =12
        'DataGrid1.Columns(cst_券別).Visible = False
        'If sm.UserInfo.TPlanID = "12" Then DataGrid1.Columns(cst_券別).Visible = True

        '報名管道(職前計畫顯示)
        'If blnP0 Then DataGrid1.Columns(cst_券別).Visible = True
        'If blnP0 Then DataGrid1.Columns(cst_甄試加分).Visible = True

        '就服單位協助報名
        'If sm.UserInfo.Years >= 2013 AndAlso TIMS.Cst_TPlanID0237AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    DataGrid1.Columns(cst_券別).Visible = True
        '    DataGrid1.Columns(cst_甄試加分).Visible = True 'DataGrid1.Columns(cst_免試).Visible = True
        'End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        labmsg219.Text = "" 'cst_msg219
        If Not IsPostBack Then
            DataGridTable.Visible = False
            msg.Text = ""
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Call Button5_Click(sender, e)
            End If
            TIMS.Tooltip(Button2, cst_Button2tip)
            Button1.Attributes("onclick") = "javascript:return search();"
            Button2.Attributes("onclick") = "Grade();return false;" '試算
            Button3.Attributes("onclick") = "javascript:return chkdata();"
            Button4.Attributes("onclick") = "javascript:return search();"
            Button7.Attributes("onclick") = "javascript:return search();"
        End If

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

    '取得成績計算比例,ItemVarFlag(0=>全計畫設定, 1=>只單一班設定)
    'Function getItemVar23(ByVal sqlconn As SqlConnection) As Boolean
    '    Dim rst As Boolean = True '正常為True 異常為False '並顯示訊息

    '    ItemVar1.Value = "" 'TIMS.ClearSQM(ItemVar1.Value)
    '    ItemVar2.Value = "" 'TIMS.ClearSQM(ItemVar2.Value)
    '    DataGridTable.Visible = False
    '    ArgRole.Text = ""

    '    Call TIMS.OpenDbConn(sqlconn)
    '    '查詢開放設定
    '    Dim gvid23 As String = TIMS.GetGlobalVar(Me, "23", "1", objconn)

    '    If gvid23 = "Y" Then
    '        '開放設定,取機構設定資料
    '        'sql = "" & vbCrLf
    '        Dim sql As String = ""
    '        sql = "" & vbCrLf
    '        sql &= " WITH WC1 AS (" & vbCrLf
    '        sql &= "   SELECT a.orgid" & vbCrLf
    '        sql &= "   FROM auth_relship a" & vbCrLf
    '        sql &= "   JOIN org_orginfo b ON b.orgid = a.orgid" & vbCrLf
    '        sql &= "   WHERE a.rid = @rid" & vbCrLf
    '        sql &= " )" & vbCrLf
    '        sql &= " SELECT writeresult ,oralresult ,ocid" & vbCrLf
    '        sql &= " FROM org_writeoral" & vbCrLf
    '        sql &= " WHERE 1=1" & vbCrLf
    '        sql &= "    AND planid = @planid" & vbCrLf
    '        sql &= "    AND orgid IN (SELECT ORGID FROM WC1)" & vbCrLf
    '        sql &= " ORDER BY ocid DESC" & vbCrLf
    '        Dim sCmd As New SqlCommand(sql, objconn)
    '        Dim dt As New DataTable
    '        With sCmd
    '            .Parameters.Clear()
    '            .Parameters.Add("planid", SqlDbType.VarChar).Value = sm.UserInfo.PlanID
    '            .Parameters.Add("rid", SqlDbType.VarChar).Value = RIDValue.Value
    '            'dt.Load(.ExecuteReader())
    '            dt = DbAccess.GetDataTable(sCmd.CommandText, objconn, sCmd.Parameters)
    '        End With
    '        If dt.Rows.Count = 0 Then
    '            rst = False
    '            Common.MessageBox(Me, "系統尚未設定筆試與口試的參數,請聯絡系統管理員")
    '            Return rst
    '        End If
    '        If dt.Rows.Count > 0 Then
    '            Dim dr As DataRow = dt.Rows(0)
    '            ItemVar1.Value = Convert.ToString(dr("writeresult"))
    '            ItemVar2.Value = Convert.ToString(dr("oralresult"))
    '        End If
    '    Else
    '        '未開放設定,取參數設定資訊
    '        ItemVar1.Value = TIMS.GetGlobalVar(Me, "2", "1", objconn)
    '        ItemVar2.Value = TIMS.GetGlobalVar(Me, "2", "2", objconn)
    '    End If
    '    ItemVar1.Value = TIMS.ClearSQM(ItemVar1.Value)
    '    ItemVar2.Value = TIMS.ClearSQM(ItemVar2.Value)
    '    If ItemVar1.Value = "" Then
    '        rst = False
    '        Common.MessageBox(Me, "系統尚未設定筆試的參數,請聯絡系統管理員!!")
    '        Return rst
    '    End If
    '    If ItemVar2.Value = "" Then
    '        rst = False
    '        Common.MessageBox(Me, "系統尚未設定口試的參數,請聯絡系統管理員!!")
    '        Return rst
    '    End If
    '    'DataGridTable.Visible = True
    '    'ArgRole.Text = "(筆試*" & ItemVar1.Value & "%)+(口試*" & ItemVar2.Value & "%)=總成績"
    '    Return rst
    'End Function

    '查詢功能 SQL
    Function QueryData1() As DataTable
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)

        Dim parms As New Hashtable() From {{"OCID1", OCIDValue1.Value}}

        Dim sql As String = ""
        sql &= " SELECT a.NAME" & vbCrLf '加else才不會造成name結果變null （oracle-->sqlsvr的語法問題）
        sql &= " ,a.IDNO" & vbCrLf
        sql &= " ,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK" & vbCrLf
        sql &= " ,b.SETID" & vbCrLf
        sql &= " ,CONVERT(varchar, b.EnterDate, 111) EnterDate" & vbCrLf
        sql &= " ,b.SerNum" & vbCrLf
        sql &= " ,b.OCID1" & vbCrLf
        sql &= " ,b.ExamNo" & vbCrLf
        sql &= " ,b.RelEnterDate" & vbCrLf
        sql &= " ,ISNULL(b.WriteResult,-1) WriteResult" & vbCrLf
        sql &= " ,ISNULL(b.OralResult,-1) OralResult" & vbCrLf
        sql &= " ,ISNULL(b.TotalResult,-1) TotalResult" & vbCrLf
        'sql &= " ,EXAMPLUS" & vbCrLf
        'sql &= " ,EIDENTITYID" & vbCrLf
        sql &= " ,b.TRNDMode" & vbCrLf
        sql &= " ,b.TRNDType" & vbCrLf
        'sql += "       ,b.NotExam" & vbCrLf '免試
        'IDENTITYID[身分別]-TYPE_EIdentityID
        'sql &= " ,b.EXAMPLUS" & vbCrLf '甄試加分(加權3%) EXAMPLUS
        'sql &= " ,b.EIDENTITYID" & vbCrLf
        '就服單位協助報名排前面
        sql &= " ,CASE WHEN IsNull(b.EnterPath,' ') ='W' THEN 1 ELSE 2 END WSort" & vbCrLf
        'sql &= "  ,CASE WHEN g.IDNO IS NOT NULL THEN 'Y' END GOVKILL" & vbCrLf
        sql &= " ,dbo.FN_GET_GETTRAIN3(b.OCID1) GETTRAIN3" & vbCrLf 'add 甄試方式 筆試/口試填寫權限
        sql &= " FROM STUD_ENTERTYPE b WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN STUD_ENTERTEMP a WITH(NOLOCK) ON a.SETID = b.SETID" & vbCrLf
        'sql &= " LEFT JOIN WC1G g ON g.IDNO = a.IDNO" & vbCrLf
        sql &= " WHERE b.OCID1=@OCID1 AND b.CCLID IS NULL" & vbCrLf

        Select Case rblEnterPathW.SelectedValue
            Case "Y" '是 就服單位協助報名
                'sql &= " AND ISNULL(b.EnterPath,' ') = '" & TIMS.cst_EnterPathW & "'" & vbCrLf
                sql &= " AND ISNULL(b.EnterPath,' ') = @EnterPath" & vbCrLf
                parms.Add("EnterPath", TIMS.cst_EnterPathW)
            Case "N" '不是 就服單位協助報名
                sql &= " AND ISNULL(b.EnterPath,' ') != '" & TIMS.cst_EnterPathW & "'" & vbCrLf
                sql &= " AND ISNULL(b.EnterPath,' ') != @EnterPath" & vbCrLf
                parms.Add("EnterPath", TIMS.cst_EnterPathW)
        End Select
        sql &= " ORDER BY WSort,b.ExamNo" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        Return dt
    End Function

    ''' <summary>
    ''' 查詢 班級申請 甄試方式(筆試/口試)設定結果 
    ''' </summary>
    ''' <returns></returns>
    Function GetTrain3() As String
        TMIDValue1.Value = TIMS.ClearSQM(TMIDValue1.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then Return ""
        Dim rtn As String = ""
        'Dim PlanID As String = sm.UserInfo.PlanID
        Dim sql As String = String.Format(" SELECT dbo.FN_GET_GETTRAIN3({0}) GETTRAIN3 ", OCIDValue1.Value)
        Dim dtExam As DataTable = DbAccess.GetDataTable(sql, objconn)
        If TIMS.dtHaveDATA(dtExam) Then rtn = Convert.ToString(dtExam.Rows(0)("GETTRAIN3"))
        Return rtn
    End Function

    Function GetIdentityID() As DataTable
        Dim parms As New Hashtable From {{"TPlanID", sm.UserInfo.TPlanID}, {"UNUSEDYEAR", sm.UserInfo.Years}}
        Dim sql As String = ""
        sql &= " SELECT IDENTITYID, NAME" & vbCrLf
        sql &= " FROM KEY_IDENTITY" & vbCrLf
        sql &= " WHERE IDENTITYID NOT IN (Select IDENTITYID FROM PLAN_IDENTITY WHERE TPlanID = @TPlanID AND ISENABLED = 'N')" & vbCrLf
        sql &= " AND (UNUSEDYEAR IS NULL OR UNUSEDYEAR > @UNUSEDYEAR)" & vbCrLf
        sql &= " ORDER BY IDENTITYID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        Return dt
    End Function

    '查詢功能
    Sub Search1()
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請重新查詢班級資料!!")
            Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請重新查詢班級資料!!")
            Exit Sub
        End If

        'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        Call ChkEsignUpStatus()

        RIDValue.Value = Convert.ToString(drCC("RID"))
        '取得成績計算比例,ItemVarFlag(0=>全計畫設定, 1=>只單一班設定)
        DataGridTable.Visible = False
        ArgRole.Text = ""
        Dim inParms As New Hashtable
        inParms.Add("RIDValue", Convert.ToString(drCC("RID")))
        inParms.Add("PlanID", Convert.ToString(drCC("PlanID")))
        inParms.Add("DistID", Convert.ToString(drCC("DistID")))
        inParms.Add("TPLANID", Convert.ToString(drCC("TPLANID")))
        Dim outParms As New Hashtable
        If Not TIMS.getItemVar23(sm, Me, objconn, inParms, outParms) Then
            Dim s_ErrorMsg1 As String = TIMS.GetMyValue2(outParms, "ErrorMsg1")
            Common.MessageBox(Me, s_ErrorMsg1)
            Exit Sub
        End If
        ItemVar1.Value = TIMS.GetMyValue2(outParms, "ItemVar1")
        ItemVar2.Value = TIMS.GetMyValue2(outParms, "ItemVar2")
        ArgRole.Text = TIMS.GetMyValue2(outParms, "ArgRole")

        'Dim exam As DataTable = GetTrain3()
        'Dim type As String
        'If exam.Rows(0)("GetTrain3").ToString().Contains("2") Then type = exam.Rows(0)("GetTrain3").ToString()
        Dim dt As DataTable = QueryData1()
        'dt = QueryData1()
        If ViewState("sort") Is Nothing Then ViewState("sort") = "ExamNo"

        Hid_OCID1.Value = ""
        Hid_GETTRAIN3.Value = ""

        DataGridTable.Visible = False
        msg.Text = "查無資料!"

        If TIMS.dtHaveDATA(dt) Then
            'For Each dr1 As DataRow In dt.Rows
            '    Dim flagPS1 As Boolean = TIMS.Chk_TICKETPS1(objconn, dr1("OCID1"), dr1("idno"))
            '    If flagPS1 Then
            '        dr1("NAME") = "- " & Convert.ToString(dr1("NAME"))
            '        dr1("NAME") = "Y"
            '    End If
            'Next
            'dt.AcceptChanges()

            Hid_OCID1.Value = OCIDValue1.Value
            Hid_GETTRAIN3.Value = Convert.ToString(dt.Rows(0)("GETTRAIN3")) '甄試方式設定結果

            DataGridTable.Visible = True
            msg.Text = ""

            dt.DefaultView.Sort = ViewState("sort")
            DataGrid1.DataSource = dt
            DataGrid1.DataKeyField = "SETID"
            DataGrid1.DataBind()

            If sm.UserInfo.LID > 1 AndAlso RIDValue.Value <> sm.UserInfo.RID Then
                Button2.Enabled = False
                Button3.Enabled = False
                TIMS.Tooltip(Button2, "使用者 無此權限")
                TIMS.Tooltip(Button2, "使用者 無此權限")
            End If
        End If
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Hid_OCID1.Value = ""
        Call Search1()
    End Sub

    '儲存 STUD_ENTERTYPE
    Sub SaveData1()
        Hid_GETTRAIN3.Value = GetTrain3()
        Dim blFlag2 As Boolean = Hid_GETTRAIN3.Value.Contains("2") '需筆試
        Dim blFlag3 As Boolean = Hid_GETTRAIN3.Value.Contains("3") '需口試

        'Dim s_log1 As String = ""
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim grade1 As TextBox = eItem.Cells(cst_筆試成績).Controls(1)
            Dim grade2 As TextBox = eItem.Cells(cst_口試成績).Controls(1)
            Dim grade3 As TextBox = eItem.Cells(cst_總成績).Controls(1)
            grade1.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(grade1.Text))
            grade2.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(grade2.Text))
            grade3.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(grade3.Text))
            's_log1 = String.Format("grade1:{0},grade1:{1},grade1:{2}", grade1.Text, grade2.Text, grade3.Text)
            'TIMS.writeLog(Me, s_log1)

            Dim Hid_SETID As HtmlInputHidden = eItem.FindControl("Hid_SETID")
            Dim Hid_EnterDate As HtmlInputHidden = eItem.FindControl("Hid_EnterDate")
            Dim Hid_SerNum As HtmlInputHidden = eItem.FindControl("Hid_SerNum")
            Dim SETID_V As String = Hid_SETID.Value
            Dim EnterDate_V As String = Hid_EnterDate.Value
            Dim SerNum_V As String = Hid_SerNum.Value
            's_log1 = String.Format("SETID_V:{0},EnterDate_V:{1},SerNum_V:{2}", SETID_V, EnterDate_V, SerNum_V)
            'TIMS.writeLog(Me, s_log1)

            '甄試加分(加權3%) EXAMPLUS 'Dim EXAMPLUS As HtmlInputCheckBox = eItem.FindControl("EXAMPLUS")
            '身分別            'Dim TYPE_EIdentityID As DropDownList = eItem.FindControl("TYPE_EIdentityID")
            'Dim Hid_IdentityID As HiddenField = eItem.FindControl("Hid_IdentityID")
            Dim flag_can_save As Boolean = False '要有輸入欄位啟用，即可儲存
            If grade1.Enabled OrElse grade2.Enabled Then flag_can_save = True
            If flag_can_save Then
                'Dim dt As DataTable = Nothing
                Dim pms1 As New Hashtable From {{"SETID", SETID_V}, {"EnterDate", TIMS.Cdate2(EnterDate_V)}, {"SerNum", SerNum_V}}
                Dim sql As String = " SELECT 1 FROM STUD_ENTERTYPE WITH(NOLOCK) WHERE SETID=@SETID AND EnterDate =@EnterDate AND SerNum =@SerNum"
                Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, pms1)
                If TIMS.dtHaveDATA(dt) Then
                    Dim dr As DataRow = dt.Rows(0)
                    If IsNumeric(grade1.Text) Then grade1.Text = CDbl(grade1.Text)
                    If IsNumeric(grade2.Text) Then grade2.Text = CDbl(grade2.Text)
                    If IsNumeric(grade3.Text) Then grade3.Text = CDbl(grade3.Text)
                    'If IsNumeric(grade1.Text) Then dr("WriteResult") = Val(grade1.Text) Else dr("WriteResult") = -1
                    'If IsNumeric(grade2.Text) Then dr("OralResult") = Val(grade2.Text) Else dr("OralResult") = -1

                    '2018-09-14 免試的甄試方式成績欄位全改存成-1
                    Dim i_WriteResult As Double = If(IsNumeric(grade1.Text) AndAlso blFlag2, Val(grade1.Text), -1)
                    Dim i_OralResult As Double = If(IsNumeric(grade2.Text) AndAlso blFlag3, Val(grade2.Text), -1)
                    Dim i_TotalResult As Double = If(IsNumeric(grade3.Text), Val(grade3.Text), -1)

                    If (blFlag2 AndAlso Val(grade1.Text) = -1) OrElse (blFlag3 AndAlso Val(grade2.Text) = -1) Then
                        '只要需筆試但註記缺考（填-1）or 需口試但註記缺考（填-1）==>不計算總分
                        i_TotalResult = -1
                    End If
                    'dr("EXAMPLUS") = Convert.DBNull 'IIf(EXAMPLUS.Checked, 1, 0) '甄試加分(加權3%) EXAMPLUS
                    '預設為空 NULL (若不為空就儲存)
                    'dr("EIdentityID") = Convert.DBNull ' TIMS.GetValue1(TYPE_EIdentityID.SelectedValue)
                    'dr("ModifyAcct") = sm.UserInfo.UserID
                    'dr("ModifyDate") = Now
                    'dr("NotExam") = NotExam.Checked

                    Dim ParmsU As New Hashtable From {
                        {"WRITERESULT", i_WriteResult},
                        {"ORALRESULT", i_OralResult},
                        {"TOTALRESULT", i_TotalResult},
                        {"ModifyAcct", sm.UserInfo.UserID},
                        {"SETID", Val(SETID_V)},
                        {"ENTERDATE", TIMS.Cdate2(EnterDate_V)},
                        {"SERNUM", Val(SerNum_V)}
                    }
                    Dim uSql As String = ""
                    uSql &= " UPDATE STUD_ENTERTYPE" & vbCrLf
                    uSql &= " SET WRITERESULT = @WRITERESULT ,ORALRESULT = @ORALRESULT ,TOTALRESULT = @TOTALRESULT" & vbCrLf
                    uSql &= " ,MODIFYACCT = @MODIFYACCT ,MODIFYDATE = GETDATE()" & vbCrLf
                    uSql &= " WHERE  SETID = @SETID AND ENTERDATE = @ENTERDATE AND SERNUM = @SERNUM" & vbCrLf
                    DbAccess.ExecuteNonQuery(uSql, objconn, ParmsU)
                End If
                'If dt.Rows.Count <> 0 Then ThenEnd IfDbAccess.UpdateDataTable(dt, da)
            End If
        Next

        'conn.Close()
        'Button1_Click(sender, e)
        Call Search1()
        'Common.RespWrite(Me, "<script language=javascript>window.alert('資料更新成功!');</script>")
        Page.RegisterStartupScript("SaveData1", "<script>alert('資料更新成功');</script>")
    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True

        If Hid_OCID1.Value = "" Then
            Errmsg &= "請重新查詢班級資料!!" & vbCrLf
            Return False
        End If
        Dim drOC As DataRow = TIMS.GetOCIDDate(Hid_OCID1.Value, objconn)
        If drOC Is Nothing Then
            Errmsg &= "請重新查詢班級資料!!" & vbCrLf
            Return False
        End If

        'https://jira.turbotech.com.tw/browse/TIMSC-161
        '非系統管理者
        Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1)
        If Not flagS1 Then
            Dim dtArc As DataTable '暫時權限Table
            dtArc = TIMS.Get_Auth_REndClass(Me, objconn)
            If TIMS.ChkIsEndDate(OCIDValue1.Value, TIMS.cst_FunID_甄試成績登錄, dtArc) Then
                '過了使用期限 True(不可使用)   False(可使用)
                If $"{drOC("InputOK")}" <> "Y" Then
                    Errmsg &= "甄試成績登錄已過開訓日，開訓日後鎖定功能填寫。" & vbCrLf
                    Return False
                End If
            End If
        End If

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    '儲存
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請重新查詢班級資料!!")
            Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請重新查詢班級資料!!")
            Exit Sub
        End If

        '先檢查是不是全部都是數字'rst = False  'Return rst '異常
        If Not Chknum() Then Exit Sub



        '(檢查 / 試算) '試算看看 有問題為false 正常為true.
        Dim iPen As Integer = 0
        Dim iTalk As Integer = 0
        If Not sUtl_Action1(iPen, iTalk) Then Exit Sub

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Hid_OCID1.Value = TIMS.ClearSQM(Hid_OCID1.Value)
        If Hid_OCID1.Value = "" Then
            Common.MessageBox(Me, "請重新查詢班級資料!!")
            Exit Sub
        End If
        If Hid_OCID1.Value <> OCIDValue1.Value Then
            Common.MessageBox(Me, "請重新查詢班級資料!!")
            Exit Sub
        End If

        Call SaveData1() '儲存
    End Sub

    '檢查是否尚有e網報名未審學員
    Sub ChkEsignUpStatus()
        Dim dt As DataTable
        Dim parms As New Hashtable From {{"OCID1", OCIDValue1.Value}}
        'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        Dim SSql As String = ""
        SSql &= " SELECT e.Name ,e.IDNO ,dbo.FN_GET_MASK1(e.IDNO) IDNO_MK" & vbCrLf
        SSql &= " ,a.eSETID ,a.RelEnterDate" & vbCrLf
        SSql &= " ,a.signUpStatus" & vbCrLf
        SSql &= " ,b.OCID ,d.OrgName" & vbCrLf
        SSql &= " ,b.ClassCName" & vbCrLf
        SSql &= " FROM STUD_ENTERTYPE2 a WITH(NOLOCK)" & vbCrLf
        SSql &= " JOIN STUD_ENTERTEMP2 e WITH(NOLOCK) ON a.eSETID=e.eSETID" & vbCrLf
        SSql &= " JOIN CLASS_CLASSINFO b WITH(NOLOCK) ON a.OCID1=b.OCID" & vbCrLf
        SSql &= " JOIN AUTH_RELSHIP c WITH(NOLOCK) ON b.RID=c.RID" & vbCrLf
        SSql &= " JOIN ORG_ORGINFO d WITH(NOLOCK) ON c.OrgID=d.Orgid" & vbCrLf
        SSql &= " WHERE a.signUpStatus=0 AND a.OCID1=@OCID1" & vbCrLf
        SSql &= " ORDER BY a.RelEnterDate" & vbCrLf
        dt = DbAccess.GetDataTable(SSql, objconn, parms)

        LinkButton1.Visible = False
        DataGrid2.Visible = False
        Button2.Visible = True
        Button3.Visible = True
        Button4.Visible = True

        If TIMS.dtNODATA(dt) Then Return
        'If dt.Rows.Count > 0 Then 'End If

        LinkButton1.Visible = True
        DataGrid2.Visible = True
        Button2.Visible = False
        Button3.Visible = False
        Button4.Visible = False

        DataGrid2.DataSource = dt
        DataGrid2.DataKeyField = "eSETID"
        DataGrid2.DataBind()
        Common.MessageBox(Me.Page, "目前尚有e網報名學員尚未審核，需先將e網報名的所有學員審核完成後，始能登錄甄試成績！！")

    End Sub

    '檢查數字'先檢查是不是全部都是數字
    Function Chknum() As Boolean
        Dim Rst As Boolean = True 'False:有問題數字 True:數字ok
        'Dim i, j, k 'Dim num As String = "0123456789" 'chknum = False

        For Each eItem As DataGridItem In DataGrid1.Items
            Dim grade1 As TextBox = eItem.Cells(cst_筆試成績).Controls(1)
            Dim grade2 As TextBox = eItem.Cells(cst_口試成績).Controls(1)
            Dim grade3 As TextBox = eItem.Cells(cst_總成績).Controls(1)
            grade1.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(grade1.Text))
            grade2.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(grade2.Text))
            grade3.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(grade3.Text))
            'Dim Hid_SETID As HtmlInputHidden = eItem.FindControl("Hid_SETID")
            'Dim Hid_EnterDate As HtmlInputHidden = eItem.FindControl("Hid_EnterDate")
            'Dim Hid_SerNum As HtmlInputHidden = eItem.FindControl("Hid_SerNum")
            'Dim SETID_V As String = Hid_SETID.Value
            'Dim EnterDate_V As String = Hid_EnterDate.Value
            'Dim SerNum_V As String = Hid_SerNum.Value

            '甄試加分(加權3%) EXAMPLUS
            'Dim EXAMPLUS As HtmlInputCheckBox = eItem.FindControl("EXAMPLUS")

            '身分別
            'Dim TYPE_EIdentityID As DropDownList = eItem.FindControl("TYPE_EIdentityID")
            'Dim Hid_IdentityID As HiddenField = eItem.FindControl("Hid_IdentityID")

            If grade1.Text <> "" Then
                If IsNumeric(grade1.Text) Then
                    Try
                        grade1.Text = CDbl(grade1.Text)
                        If CInt(grade1.Text) > 100 Then
                            Common.RespWrite(Me, "<script language=javascript>window.alert('筆試成績不能大於100!');</script>")
                            Rst = False
                            Exit For
                            'ElseIf grade1.Text < 0 Then
                            'Common.RespWrite(Me, "<script language=javascript>window.alert('筆成績不能小於0!');</script>")
                            'Return False
                        End If
                    Catch ex As Exception
                        Common.RespWrite(Me, "<script language=javascript>window.alert('筆試成績輸入格式有誤，不為數字!');</script>")
                        Rst = False
                        Exit For
                    End Try
                Else
                    Common.RespWrite(Me, "<script language=javascript>window.alert('筆試成績輸入格式有誤，不為數字!');</script>")
                    Rst = False
                    Exit For
                    'Return False
                End If
            End If

            If grade2.Text <> "" Then
                If IsNumeric(grade2.Text) Then
                    Try
                        grade2.Text = CDbl(grade2.Text)
                        If CInt(grade2.Text) > 100 Then
                            Common.RespWrite(Me, "<script language=javascript>window.alert('口試成績不能大於100!');</script>")
                            Rst = False
                            Exit For
                            'Return False
                            'ElseIf grade2.Text < 0 Then
                            'Common.RespWrite(Me, "<script language=javascript>window.alert('口試成績不能小於0!');</script>")
                            'Return False
                        End If
                    Catch ex As Exception
                        Common.RespWrite(Me, "<script language=javascript>window.alert('口試成績輸入格式有誤，不為數字!');</script>")
                        Rst = False
                        Exit For
                    End Try
                Else
                    Common.RespWrite(Me, "<script language=javascript>window.alert('口試成績輸入格式有誤，不為數字!');</script>")
                    Rst = False
                    Exit For
                    'Return False
                End If
            End If

            If grade3.Text <> "" Then
                If IsNumeric(grade3.Text) Then
                    Try
                        grade3.Text = CDbl(grade3.Text)
                        If CInt(grade3.Text) > 100 Then
                            Common.RespWrite(Me, "<script language=javascript>window.alert('總成績不能大於100!');</script>")
                            Rst = False
                            Exit For
                            'Return False
                            'ElseIf grade2.Text < 0 Then
                            'Common.RespWrite(Me, "<script language=javascript>window.alert('口試成績不能小於0!');</script>")
                            'Return False
                        End If
                    Catch ex As Exception
                        Common.RespWrite(Me, "<script language=javascript>window.alert('總成績輸入格式有誤，不為數字!');</script>")
                        Rst = False
                        Exit For
                    End Try
                Else
                    Common.RespWrite(Me, "<script language=javascript>window.alert('總成績輸入格式有誤，不為數字!');</script>")
                    Rst = False
                    Exit For
                    'Return False
                End If
            End If

            'If EXAMPLUS.Checked AndAlso TYPE_EIdentityID.SelectedValue = "" Then
            '    Common.MessageBox(Me, "有勾選「加權3%」者，需必選「身分別」欄!")
            '    Rst = False '異常
            '    Exit For '異常
            'End If
        Next

        Return Rst
        'chknum = True
        'Return True
    End Function

    '(檢查 / 試算) '試算看看 有問題為false 正常為true.
    Function sUtl_Action1(ByRef iPen As Integer, ByRef iTalk As Integer) As Boolean
        Dim rst As Boolean = True 'True :正常結束 False:異常，離開。

        '先取出計畫代碼中的訓練計畫編號----------------------Start
        Dim pms1 As New Hashtable From {{"PlanID", sm.UserInfo.PlanID}}
        Dim sql As String = "SELECT TPLANID FROM ID_PLAN WHERE PlanID=@PlanID"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, pms1)
        If dr Is Nothing Then
            Common.MessageBox(Me, "該計畫代碼異常，請聯絡 系統管理者 確認系統參數!")
            'Exit Function rst = False
            Return False '異常
        End If
        '先取出計畫代碼中的訓練計畫編號----------------------End

        'Dim drCC As DataRow = TIMS.GetOCIDDate(Hid_OCID1.Value, objconn)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請重新查詢班級資料!!")
            Return False '異常 Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請重新查詢班級資料!!")
            Return False '異常  Exit Sub
        End If

        DataGridTable.Visible = False
        ArgRole.Text = ""
        Dim inParms As New Hashtable
        'inParms.Add("RIDValue", RIDValue.Value)
        inParms.Add("RIDValue", Convert.ToString(drCC("RID")))
        inParms.Add("PlanID", Convert.ToString(drCC("PlanID")))
        inParms.Add("DistID", Convert.ToString(drCC("DistID")))
        inParms.Add("TPLANID", Convert.ToString(drCC("TPLANID")))
        Dim outParms As New Hashtable
        If Not TIMS.getItemVar23(sm, Me, objconn, inParms, outParms) Then
            Dim s_ErrorMsg1 As String = TIMS.GetMyValue2(outParms, "ErrorMsg1")
            'Common.MessageBox(Me, s_ErrorMsg1)
            Common.MessageBox(Me, "尚未設定系統參數「筆試與口試成績比例!」，" & vbCrLf & "請聯絡分署的人設定本計劃的系統參數!")
            Return False '異常'  Exit Sub
        End If
        ItemVar1.Value = TIMS.GetMyValue2(outParms, "ItemVar1")
        ItemVar2.Value = TIMS.GetMyValue2(outParms, "ItemVar2")
        ArgRole.Text = TIMS.GetMyValue2(outParms, "ArgRole")

        'If Not getItemVar23(objconn) Then
        '    Common.MessageBox(Me, "尚未設定系統參數「筆試與口試成績比例!」，" & vbCrLf & "請聯絡分署的人設定本計劃的系統參數!")
        '    'Exit Function rst = False
        '    Return False '異常
        'End If
        iPen = Val(ItemVar1.Value)
        iTalk = Val(ItemVar2.Value)
        '檢查系統參數是否有設定------------------------------End

        Try
            '試著試算一下。
            For Each eItem As DataGridItem In DataGrid1.Items
                Dim grade1 As TextBox = eItem.Cells(cst_筆試成績).Controls(1)
                Dim grade2 As TextBox = eItem.Cells(cst_口試成績).Controls(1)
                Dim grade3 As TextBox = eItem.Cells(cst_總成績).Controls(1)
                grade1.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(grade1.Text))
                grade2.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(grade2.Text))
                grade3.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(grade3.Text))
                '甄試加分(加權3%) EXAMPLUS
                'Dim EXAMPLUS As HtmlInputCheckBox = eItem.FindControl("EXAMPLUS")
                '身分別
                'Dim TYPE_EIdentityID As DropDownList = eItem.FindControl("TYPE_EIdentityID")
                'Dim Hid_IdentityID As HiddenField = eItem.FindControl("Hid_IdentityID")

                Dim iGrade1 As Double = 0 '任何異常都是0
                Dim iGrade2 As Double = 0 '任何異常都是0
                Dim iGrade3 As Double = 0 '任何異常都是0
                If grade1.Text <> "" AndAlso Val(grade1.Text) >= 0 Then iGrade1 = Val(grade1.Text)
                If grade2.Text <> "" AndAlso Val(grade2.Text) >= 0 Then iGrade2 = Val(grade2.Text)
                Dim flag_can_save As Boolean = False '要有輸入欄位啟用，即可儲存
                If grade1.Enabled OrElse grade2.Enabled Then flag_can_save = True
                If flag_can_save Then
                    If iGrade1 = -1 OrElse iGrade2 = -1 Then
                        iGrade3 = -1
                    Else
                        'If EXAMPLUS.Checked AndAlso TYPE_EIdentityID.SelectedValue = "" Then
                        '    Common.MessageBox(Me, "有勾選「加權3%」者，需必選「身分別」欄!")
                        '    Return False '異常
                        'End If
                        Dim iNum As Double = 0
                        '甄試加分(加權3%) EXAMPLUS
                        'If EXAMPLUS.Checked Then
                        '    iNum = ((iGrade1 * iPen / 100) + (iGrade2 * iTalk / 100)) * 1.03
                        'Else
                        '    iNum = ((iGrade1 * iPen / 100) + (iGrade2 * iTalk / 100))
                        'End If
                        iNum = ((iGrade1 * iPen / 100) + (iGrade2 * iTalk / 100))
                        iNum = TIMS.ROUND(iNum, 1)
                        If iNum > 100 Then iNum = 100
                        iGrade3 = iNum
                    End If
                    grade3.Text = iGrade3
                End If
            Next
        Catch ex As Exception
            'Common.MessageBox(Me, ex.ToString)
            Common.MessageBox(Me, "試算未正常結束，請檢查輸入值!")
            'rst = False
            Dim strErrmsg As String = "" & vbCrLf
            'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg &= "ex.ToString:" & ex.ToString & vbCrLf
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)

            Return False '異常
        End Try

        Return rst
    End Function

    '匯出(匯入甄試成績)
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim dt As DataTable = QueryData1()
        Call Creattable(dt)
    End Sub

    Sub SImportFile1(ByRef FullFileName1 As String)
        '上傳檔案 'File1.PostedFile.SaveAs(Server.MapPath("~/SD/03/Temp/" & sMyFileName)) '上傳檔案
        File1.PostedFile.SaveAs(FullFileName1)

        Dim dt_xls As DataTable = Nothing
        Dim Reason As String = "" '儲存錯誤的原因
        '取得內容
        If (flag_File1_xls) Then
            Const cst_FirstCol1 As String = "編號"
            dt_xls = TIMS.GetDataTable_XlsFile(FullFileName1, "", Reason, cst_FirstCol1)
            If Reason <> "" Then
                Common.MessageBox(Me, "無法匯入!!" & Reason)
                Exit Sub
            End If
        End If
        If (flag_File1_ods) Then dt_xls = TIMS.GetDataTable_ODSFile(FullFileName1)
        If (flag_File1_csv) Then dt_xls = TIMS.GetDataTable_CSVFile(FullFileName1)

        '刪除檔案 'IO.File.Delete(FullFileName1)
        TIMS.MyFileDelete(FullFileName1)

        If TIMS.dtNODATA(dt_xls) Then
            Common.MessageBox(Me, "資料有誤，故無法匯入，請修正匯入檔案，謝謝")
            Exit Sub
        End If

        Const cst_缺考 As String = "缺考"
        Dim sMsgBox As String = ""
        'add 匯入時加入免匯入筆試/口試成績的提示訊息
        Dim strGetTrain3 As String = GetTrain3()
        Dim blFlag2 As Boolean = strGetTrain3.Contains("2") '需筆試
        Dim blFlag3 As Boolean = strGetTrain3.Contains("3") '需口試
        '取得身分別資料
        Dim identityDt As DataTable = GetIdentityID()

        Dim iRowIndex As Integer = 1
        For Each drCOL1 As DataRow In dt_xls.Rows
            Dim ExamNo As String = ""
            Dim Name As String = ""
            Dim IDNO As String = ""
            Dim RelEnterDate As String = ""
            Dim WriteResult As String = ""
            Dim OralResult As String = ""
            Dim TotalResult As String = ""

            Try
                ExamNo = drCOL1(cst_a准考證號碼).ToString
                'sernum = colArray(cst_a編號).ToString
                Name = drCOL1(cst_a姓名).ToString
                IDNO = drCOL1(cst_a身分證號碼).ToString
                RelEnterDate = drCOL1(cst_a報名日期).ToString
                WriteResult = If(drCOL1(cst_a筆試成績).ToString.Length <> 0, drCOL1(cst_a筆試成績).ToString, cst_缺考)
                OralResult = If(drCOL1(cst_a口試成績).ToString.Length <> 0, drCOL1(cst_a口試成績).ToString, cst_缺考)
                'Examplus = Convert.ToString(colArray(cst_a加權))
                'EIdentityID = Convert.ToString(colArray(cst_a身分別代碼))
                TotalResult = If(drCOL1(cst_a總成績).ToString.Length <> 0, drCOL1(cst_a總成績).ToString, cst_缺考)
            Catch ex As Exception
                sMsgBox += "ex: " & ex.ToString
            End Try

            If ExamNo = "" Then sMsgBox += "請檢查第 " & iRowIndex & " 筆資料(准考證號/報名序號為空)" & vbCrLf

            IDNO = TIMS.ClearSQM(IDNO)
            If IDNO <> "" Then
                Dim rqIDNO As String = IDNO
                '1:國民身分證 2:居留證 4:居留證2021
                Dim flag1 As Boolean = TIMS.CheckIDNO(rqIDNO)
                Dim flag2 As Boolean = TIMS.CheckIDNO2(rqIDNO, 2)
                Dim flag4 As Boolean = TIMS.CheckIDNO2(rqIDNO, 4)
                If Not flag1 AndAlso Not flag2 AndAlso Not flag4 Then sMsgBox += "請檢查第 " & iRowIndex & " 筆資料(身分證號或居留證號有誤)" & String.Format("[{0}]", IDNO) & vbCrLf

            Else
                sMsgBox += "請檢查第 " & iRowIndex & " 筆資料(身分證號為空)" & vbCrLf
            End If

            '2018-09-13 add 要筆試時才做成績輸入檢核
            If blFlag2 AndAlso WriteResult <> cst_缺考 Then
                If Not IsNumeric(WriteResult) Then
                    sMsgBox &= "請檢查第 " & iRowIndex & " 筆資料(筆試成績請輸入數字)" & String.Format("[{0}]", WriteResult) & vbCrLf
                Else
                    If Val(WriteResult) > 100 Then sMsgBox &= "請檢查第 " & iRowIndex & " 筆資料(筆試成績不能大於100)" & String.Format("[{0}]", WriteResult) & vbCrLf
                End If
            End If

            '2018-09-13 add 要做口試時才做成績輸入檢核
            If blFlag3 AndAlso OralResult <> cst_缺考 Then
                If Not IsNumeric(OralResult) Then
                    sMsgBox &= "請檢查第 " & iRowIndex & " 筆資料(口試成績請輸入數字)" & String.Format("[{0}]", OralResult) & vbCrLf
                Else
                    If Val(OralResult) > 100 Then sMsgBox &= "請檢查第 " & iRowIndex & " 筆資料(口試成績不能大於100)" & String.Format("[{0}]", OralResult) & vbCrLf
                End If
            End If

            'add 檢核加權3%
            'If Examplus <> "" Then
            '    If Examplus <> "Y" AndAlso Examplus <> "N" Then sMsgBox &= "請檢查第 " & RowIndex & " 筆資料(是否加權3%請輸入Y或N)" & vbCrLf
            'End If

            'add 檢核身分別
            'If Examplus = "Y" Then
            '    If EIdentityID.Trim() = "" Then
            '        sMsgBox &= "請檢查第 " & RowIndex & " 筆資料(有「加權3%」者，請輸入「身分別代碼」)" & vbCrLf
            '    ElseIf identityDt.Select("IDENTITYID='" & EIdentityID.PadLeft(2, "0") & "'").Length = 0 Then
            '        sMsgBox &= "請檢查第 " & RowIndex & " 筆資料(「身分別代碼」輸入錯誤)" & vbCrLf
            '    End If
            'ElseIf EIdentityID.Trim() <> "" Then
            '    If identityDt.Select("IDENTITYID='" & EIdentityID.PadLeft(2, "0") & "'").Length = 0 Then sMsgBox &= "請檢查第 " & RowIndex & " 筆資料(「身分別代碼」輸入錯誤)" & vbCrLf
            'End If

            If TotalResult <> cst_缺考 Then
                If Not IsNumeric(TotalResult) Then
                    sMsgBox &= "請檢查第 " & iRowIndex & " 筆資料(總成績請輸入數字)" & String.Format("[{0}]", TotalResult) & vbCrLf
                Else
                    If Val(TotalResult) > 100 Then sMsgBox &= "請檢查第 " & iRowIndex & " 筆資料(總成績不能大於100)" & vbCrLf
                End If
            End If

            Dim dtEnter1 As DataTable = Nothing
            If sMsgBox = "" Then
                IDNO = TIMS.ChangeIDNO(IDNO)
                Dim parms As New Hashtable From {{"OCID1", OCIDValue1.Value}, {"ExamNo", ExamNo}, {"IDNO", IDNO}} 'parms.Clear()
                Dim sql1 As String = ""
                sql1 &= " SELECT se1.SETID ,CONVERT(varchar, se1.EnterDate, 111) EnterDate ,se1.SerNum "
                sql1 &= " FROM Stud_EnterType se1 WITH(NOLOCK)"
                sql1 &= " JOIN Stud_EnterTemp se WITH(NOLOCK) ON se.SETID=se1.SETID "
                sql1 &= " WHERE se1.OCID1=@OCID1 AND se1.ExamNo=@ExamNo AND se.IDNO=@IDNO "
                If Name <> "" Then
                    sql1 &= " AND se.Name LIKE @Name"
                    parms.Add("Name", Replace(Name, "?", "") & "%")
                End If
                dtEnter1 = DbAccess.GetDataTable(sql1, objconn, parms)
                If TIMS.dtNODATA(dtEnter1) Then
                    sMsgBox &= "請檢查第 " & iRowIndex & " 筆資料(資料不正確,查無資料)" & vbCrLf
                Else
                    If dtEnter1.Rows.Count > 1 Then sMsgBox &= "請檢查第 " & iRowIndex & " 筆資料(資料有超過1筆以上，請確認)" & vbCrLf
                End If
            End If
            If sMsgBox = "" AndAlso dtEnter1 Is Nothing Then sMsgBox &= "請檢查第 " & iRowIndex & " 筆資料(資料不正確,查無資料)" & vbCrLf

            If sMsgBox = "" Then
                Dim dr1 As DataRow = dtEnter1.Rows(0) '只會有1筆資料
                Dim htSS As New Hashtable 'htSS Hashtable()
                htSS.Add("WRITERESULT", If(WriteResult = cst_缺考 Or Not blFlag2, "-1", WriteResult)) '筆試存入缺考時轉為-1，免試則直接存-1
                htSS.Add("ORALRESULT", If(OralResult = cst_缺考 Or Not blFlag3, "-1", OralResult)) '口試存入缺考時轉為-1，免試則直接存-1
                htSS.Add("TOTALRESULT", If(TotalResult = cst_缺考, "-1", TotalResult))
                'htSS.Add("EXAMPLUS", IIf(Examplus = "Y", "1", "0")) 'add 是加權3%
                'htSS.Add("EIDENTITYID", EIdentityID) 'add 身分別代碼
                htSS.Add("SETID", Convert.ToString(dr1("SETID")))
                htSS.Add("ENTERDATE", TIMS.Cdate3(Convert.ToString(dr1("ENTERDATE"))))
                htSS.Add("SERNUM", Convert.ToString(dr1("SERNUM")))
                Call UPDATE_STUD_ENTERTYPE(htSS)
            End If

            iRowIndex = iRowIndex + 1
        Next

        Dim s_msg As String = If(Not blFlag2, cst_imp_免筆試, If(Not blFlag3, cst_imp_免口試, ""))
        If sMsgBox <> "" Then
            sMsgBox = "有些資料匯入成功，但有錯誤的資料無法匯入，請檢查下列資料:" & vbCrLf & sMsgBox
            Common.MessageBox(Me, s_msg + "<br>" + sMsgBox)
            Exit Sub
        End If
        Common.MessageBox(Me, s_msg + "<br>資料匯入成功，請按查詢，查看匯入資料")

    End Sub

    ''' <summary>
    ''' 匯入(匯入甄試成績) csv
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '取得成績計算比例,ItemVarFlag(0=>全計畫設定, 1=>只單一班設定)
        'If Not getItemVar23(objconn) Then Exit Sub
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請重新查詢班級資料!!")
            Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請重新查詢班級資料!!")
            Exit Sub
        End If

        DataGridTable.Visible = False
        ArgRole.Text = ""
        Dim inParms As New Hashtable From {{"RIDValue", Convert.ToString(drCC("RID"))}, {"PlanID", Convert.ToString(drCC("PlanID"))}, {"DistID", Convert.ToString(drCC("DistID"))}, {"TPLANID", Convert.ToString(drCC("TPLANID"))}}
        Dim outParms As New Hashtable
        If Not TIMS.getItemVar23(sm, Me, objconn, inParms, outParms) Then
            Dim s_ErrorMsg1 As String = TIMS.GetMyValue2(outParms, "ErrorMsg1")
            Common.MessageBox(Me, s_ErrorMsg1)
            'Common.MessageBox(Me, "尚未設定系統參數「筆試與口試成績比例!」，" & vbCrLf & "請聯絡分署的人設定本計劃的系統參數!")
            Exit Sub
        End If
        ItemVar1.Value = TIMS.GetMyValue2(outParms, "ItemVar1")
        ItemVar2.Value = TIMS.GetMyValue2(outParms, "ItemVar2")
        ArgRole.Text = TIMS.GetMyValue2(outParms, "ArgRole")

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim Errmsg As String = ""
        If Convert.ToString(OCIDValue1.Value) = "" OrElse Not IsNumeric(OCIDValue1.Value) Then Errmsg += "職類/班級 有誤 請重新選擇" & vbCrLf
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If


        Dim sMyFileName As String = ""
        'Dim flag_File1_xls As Boolean = False
        'Dim flag_File1_ods As Boolean = False
        Dim sErrMsg As String = TIMS.ChkFile1(File1, sMyFileName, flag_File1_xls, flag_File1_ods, flag_File1_csv)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Return
        End If
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If flag_File1_xls Then
            If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, "xls") Then Return
        ElseIf flag_File1_ods Then
            If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, "ods") Then Return
        End If

        Const Cst_FileSavePath As String = "~/SD/03/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        sMyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim FullFileName1 As String = Server.MapPath(Cst_FileSavePath & sMyFileName)
        Call SImportFile1(FullFileName1)
    End Sub

    ''' <summary>
    ''' 匯出SQL
    ''' </summary>
    ''' <param name="dt"></param>
    Sub Creattable(ByVal dt As DataTable)
        If TIMS.dtNODATA(dt) Then
            Common.MessageBox(Me, "查無班級資料!!")
            Exit Sub
        End If

        Const cst_缺考 As String = "缺考"

        Dim sFileName1 As String = "Result"
        Dim strHTML As String = ""

        strHTML &= ("<div>")
        strHTML &= ("<table>")
        '建立輸出文字
        Dim ExportStr As String = ""
        '第1行
        ExportStr = "<tr>"
        ExportStr &= "<td>准考證號碼</td>" & vbTab
        ExportStr &= "<td>編號</td>" & vbTab
        ExportStr &= "<td>姓名</td>" & vbTab
        ExportStr &= "<td>身分證號碼</td>" & vbTab
        ExportStr &= "<td>報名日期</td>" & vbTab
        ExportStr &= "<td>筆試成績</td>" & vbTab
        ExportStr &= "<td>口試成績</td>" & vbTab
        'ExportStr &= "是否加權3% (Y/N)" & vbTab
        'ExportStr &= "身分別代碼" & vbTab
        ExportStr &= "<td>總成績</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= ExportStr

        '建立資料面
        For Each dr As DataRow In dt.Select("", "ExamNo")
            Dim str_NO As String = If(Convert.ToString(dr("ExamNo")).Length > 3, Right(Convert.ToString(dr("ExamNo")), 3), "---")
            Dim TXT_WriteResult As String = If(Convert.ToString(dr("WriteResult")) = "-1", cst_缺考, Convert.ToString(dr("WriteResult")))
            Dim TXT_OralResult As String = If(Convert.ToString(dr("OralResult")) = "-1", cst_缺考, Convert.ToString(dr("OralResult")))
            'Dim TXT_EXAMPLUS As String = If(Convert.ToString(dr("EXAMPLUS")) = "1", "Y", "N") '2018-09-14 add 是否加權3%
            'Dim TXT_EIDENTITYID As String = Convert.ToString(dr("EIDENTITYID")) '2018-09-14 add 身分別代碼
            Dim TXT_TotalResult As String = If(Convert.ToString(dr("TotalResult")) = "-1", cst_缺考, Convert.ToString(dr("TotalResult")))

            ExportStr = "<tr>"
            ExportStr &= String.Format("<td>{0}</td>", dr("ExamNo")) & vbTab '准考證號碼
            ExportStr &= String.Format("<td>{0}</td>", str_NO) & vbTab '編號
            ExportStr &= String.Format("<td>{0}</td>", dr("Name")) & vbTab '姓名
            ExportStr &= String.Format("<td>{0}</td>", dr("IDNO_MK")) & vbTab '身分證號碼
            ExportStr &= String.Format("<td>{0}</td>", Common.FormatDate(dr("RelEnterDate"))) & vbTab '報名日期
            ExportStr &= String.Format("<td>{0}</td>", TXT_WriteResult) & vbTab '筆試成績-WriteResult
            ExportStr &= String.Format("<td>{0}</td>", TXT_OralResult) & vbTab '口試成績-OralResult
            'ExportStr &= "是否加權3% (Y/N)" & vbTab'是否加權3%
            'ExportStr &= "身分別代碼" & vbTab'身分別代碼
            ExportStr &= String.Format("<td>{0}</td>", TXT_TotalResult) & vbTab '總成績
            ExportStr &= "</tr>" & vbCrLf
            strHTML &= ExportStr
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        'parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        'TIMS.CloseDbConn(objconn)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    '單一班級選擇
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGridTable.Visible = False
    End Sub

    '匯入成績(新式讀卡機-甄試成績匯入)
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "班級選擇有誤!!")
            Exit Sub
        End If
        'If OCIDValue1.Value = "" Then Exit Sub
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "班級選擇有誤!!")
            Exit Sub
        End If

        '取得成績計算比例,ItemVarFlag(0=>全計畫設定, 1=>只單一班設定)
        'If Not getItemVar23(objconn) Then Exit Sub
        DataGridTable.Visible = False
        ArgRole.Text = ""
        'inParms.Add("RIDValue", RIDValue.Value)
        Dim inParms As New Hashtable From {
            {"RIDValue", Convert.ToString(drCC("RID"))},
            {"PlanID", Convert.ToString(drCC("PlanID"))},
            {"DistID", Convert.ToString(drCC("DistID"))},
            {"TPLANID", Convert.ToString(drCC("TPLANID"))}
        }
        Dim outParms As New Hashtable
        If Not TIMS.getItemVar23(sm, Me, objconn, inParms, outParms) Then
            Dim s_ErrorMsg1 As String = TIMS.GetMyValue2(outParms, "ErrorMsg1")
            Common.MessageBox(Me, s_ErrorMsg1)
            'Common.MessageBox(Me, "尚未設定系統參數「筆試與口試成績比例!」，" & vbCrLf & "請聯絡分署的人設定本計劃的系統參數!")
            Exit Sub
        End If
        ItemVar1.Value = TIMS.GetMyValue2(outParms, "ItemVar1")
        ItemVar2.Value = TIMS.GetMyValue2(outParms, "ItemVar2")
        ArgRole.Text = TIMS.GetMyValue2(outParms, "ArgRole")

        Dim sMyFileName As String = ""
        'Dim flag_File1_xls As Boolean = False 'Dim flag_File1_ods As Boolean = False
        Dim sErrMsg As String = TIMS.ChkFile1(File2, sMyFileName, flag_File1_xls, flag_File1_ods)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Return
        End If
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If flag_File1_xls Then
            If Not TIMS.HttpCHKFile(Me, File2, MyPostedFile, "xls", 1) Then Return
        ElseIf flag_File1_ods Then
            If Not TIMS.HttpCHKFile(Me, File2, MyPostedFile, "ods", 1) Then Return
        End If

        Const Cst_FileSavePath As String = "~/SD/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File2.PostedFile.FileName).ToLower()
        sMyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim FullFileName2 As String = Server.MapPath(Cst_FileSavePath & sMyFileName)
        Call SImportFile2(FullFileName2)
        'Call Save_xls1()
    End Sub

    '檢核
    Function CheckImportData(ByVal colArray As Array, ByVal dtStudType As DataTable) As String
        'Dim aCTID As String
        Dim Reason As String = ""
        'Dim sql As String'Dim dr As DataRow'Const cst_Len As Integer = 111'Const cst_Len2 As Integer = 110

        If colArray.Length < cst_xc_Len And colArray.Length < cst_xc_Len2 Then
            Reason += "欄位對應有誤<BR>"
            Reason += "請注意欄位中是否有半形逗點<BR>"
        Else
            If colArray(4).ToString = "" Then
                Reason += "編號為必須填寫資料 <BR>"
            Else
                If Not IsNumeric(colArray(cst_xc_編號).ToString) Then
                    Reason += "編號必須為數字格式資料 <BR>"
                Else
                    If Not Len(colArray(cst_xc_編號).ToString) = 3 Then Reason += "編號必須為長度3資料 <BR>"
                End If
            End If
            If colArray(cst_xc_考生姓名).ToString = "" Then Reason += "考生姓名為必須填寫資料 <BR>"
            If colArray(cst_xc_缺考).ToString = "" Then
                Reason += "缺考欄位必須填寫資料(F/T) (有成績/缺考) <BR>"
            Else
                Select Case colArray(cst_xc_缺考).ToString
                    Case "F", "T"
                    Case Else
                        Reason += "缺考欄位必須為英文半型資料格式(F/T) (有成績/缺考)<BR>"
                End Select
            End If

            'add 判斷該班甄試方式有需做筆試時才要接續檢核
            If Hid_GETTRAIN3.Value.Contains("2") Then
                If colArray(cst_xc_成績).ToString = "" Then
                    Reason += "成績為必須填寫資料 <BR>"
                Else
                    If Not IsNumeric(colArray(cst_xc_成績).ToString) Then Reason += "成績必須為數字格式資料 <BR>"
                End If
            End If

            If colArray.Length = cst_xc_Len Then
                'add 判斷該班甄試方式有需做口試時才要接續檢核
                If Hid_GETTRAIN3.Value.Contains("3") Then
                    If colArray(cst_xc_口試).ToString = "" Then
                        'Reason += "口試為必須填寫資料 <BR>"
                    Else
                        If Not IsNumeric(colArray(cst_xc_口試).ToString) Then Reason += "口試必須為數字格式資料 <BR>"
                    End If
                End If
            End If

            If Reason = "" Then
                Dim TotalVal As Double
                TotalVal = 0
                TotalVal += (Val(colArray(cst_xc_成績).ToString) * ItemVar1.Value / 100)
                If colArray.Length = cst_xc_Len Then TotalVal += (Val(colArray(cst_xc_口試).ToString) * ItemVar2.Value / 100)
                '"(筆試*" & dr("ItemVar1") & "%)+(口試*" & dr("ItemVar2") & "%)=總成績"
                If TotalVal > 100 Then Reason += "總成績不能大於100 <BR>"
            End If
            If Reason = "" Then
                '依編號搜尋是否有資料(應為1筆資料)
                If Not dtStudType.Select("ExamNo3='" & colArray(cst_xc_編號).ToString & "'").Length = 1 Then
                    If dtStudType.Select("ExamNo3='" & colArray(cst_xc_編號).ToString & "'").Length = 0 Then
                        Reason += "依編號搜尋查無此筆資料 <BR>"
                    Else
                        Reason += "依編號搜尋查到多筆資料，請修正資料後再行匯入 <BR>"
                    End If
                End If
            End If
        End If
        Return Reason
    End Function

    '匯入成績(新式讀卡機-甄試成績匯入) xls
    Sub SImportFile2(ByRef FullFileName2 As String)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim aOCID1 As String = TIMS.ClearSQM(OCIDValue1.Value)
        If aOCID1 = "" Then
            Common.MessageBox(Me, "班別代碼有誤，請確認點選職類/班別")
            Exit Sub
        End If

        '班級學員資料
        Dim parms As New Hashtable From {{"OCID1", aOCID1}}
        Dim sql As String = ""
        sql &= " SELECT a.SETID,a.EnterDate,a.SerNum,A.ExamNo,A.WriteResult,A.OralResult,B.SETID,B.Name,b.IDNO" & vbCrLf
        sql &= " ,dbo.FN_GET_MASK1(b.IDNO) IDNO_MK" & vbCrLf
        sql &= " FROM STUD_ENTERTYPE a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN STUD_ENTERTEMP b WITH(NOLOCK) ON a.SETID=b.SETID" & vbCrLf
        sql &= " WHERE a.OCID1=@OCID1 AND a.CCLID IS NULL" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If TIMS.dtNODATA(dt) Then
            DataGridTable.Visible = False
            Common.MessageBox(Me, "此班級查無學員資料，請確認點選職類/班別")
            Exit Sub
        End If

        '上傳檔案
        File2.PostedFile.SaveAs(FullFileName2)

        Dim dt_xls As DataTable = Nothing
        Dim Reason As String = "" '儲存錯誤的原因 '取得內容
        If (flag_File1_xls) Then
            Const Cst_SheetName As String = "TMPFOX2X轉出資料"
            Const cst_FirstCol1 As String = "編號"
            dt_xls = TIMS.GetDataTable_XlsFile(FullFileName2, Cst_SheetName, Reason, cst_FirstCol1)
            If Reason <> "" Then
                Common.MessageBox(Me, "無法匯入!!" & Reason)
                Exit Sub
            End If
        End If
        If (flag_File1_ods) Then
            dt_xls = TIMS.GetDataTable_ODSFile(FullFileName2)
        End If
        '刪除檔案 'IO.File.Delete(FullFileName1)
        TIMS.MyFileDelete(FullFileName2)

        If TIMS.dtNODATA(dt_xls) Then
            Common.MessageBox(Me, "資料有誤，故無法匯入，請修正匯入檔案，謝謝")
            Exit Sub
        End If

        Dim dtS As DataTable = Nothing '班級學員資料(含編號)
        'Dim Reason As String                '儲存錯誤的原因
        Dim dtWrong As New DataTable            '儲存錯誤資料的DataTable
        Dim drWrong As DataRow

        '建立錯誤資料格式Table----------------Start
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("Name")) '考生姓名
        dtWrong.Columns.Add(New DataColumn("IDNO")) '編號
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table----------------End

        'sql += "  SUBSTRING(a.ExamNo,7,3) [ExamNo3]" & vbCrLf '★
        Dim parms3 As New Hashtable From {{"OCID1", aOCID1}}
        Dim sql3 As String = ""
        sql3 &= " SELECT substring(a.ExamNo,7,3) ExamNo3" & vbCrLf '編號。
        sql3 &= " ,a.ExamNo" & vbCrLf
        sql3 &= " ,b.Name" & vbCrLf '★
        sql3 &= " ,b.IDNO" & vbCrLf '★
        sql3 &= " ,a.SETID" & vbCrLf
        sql3 &= " ,a.EnterDate" & vbCrLf
        sql3 &= " ,a.SerNum" & vbCrLf
        sql3 &= " ,a.WriteResult" & vbCrLf
        sql3 &= " ,a.OralResult" & vbCrLf
        sql3 &= " FROM Stud_EnterType a WITH(NOLOCK)" & vbCrLf
        sql3 &= " JOIN Stud_EnterTemp b WITH(NOLOCK) ON a.SETID = b.SETID" & vbCrLf
        sql3 &= " WHERE a.OCID1=@OCID1 and CCLID IS NULL "
        dtS = DbAccess.GetDataTable(sql3, objconn, parms3)

        Dim aExamNo3 As String = ""  '編號
        Dim aName As String = ""  '姓名
        Dim aExam As String = "" '是否缺考
        Dim aWriteResult As String = "" '筆試成績
        Dim aOralResult As String = "" '口試成績
        Dim RowIndex As Integer = 0 '讀取行累計數

        'Call TIMS.OpenDbConn(tConn)
        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                'add 匯入時加入免匯入筆試/口試成績的提示訊息
                Hid_GETTRAIN3.Value = GetTrain3()
                Dim blFlag2 As Boolean = Hid_GETTRAIN3.Value.Contains("2") '需筆試
                Dim blFlag3 As Boolean = Hid_GETTRAIN3.Value.Contains("3") '需口試

                'Dim dr As DataRow
                Dim da As SqlDataAdapter = Nothing
                'dt 重設，改為存取用表格
                Dim sqlt As String = " SELECT * FROM STUD_ENTERTYPE WHERE OCID1 = '" & aOCID1 & "'" & vbCrLf
                dt = DbAccess.GetDataTable(sqlt, da, trans)

                For i As Integer = 0 To dt_xls.Rows.Count - 1
                    RowIndex = i + 1 '讀取行累計數
                    Reason = ""
                    Dim colArray As Array = dt_xls.Rows(i).ItemArray 'Split(OneRow, flag)
                    'colArray = dt_xls.Rows(i).ItemArray 'Split(OneRow, flag)
                    If Reason = "" Then Reason += CheckImportData(colArray, dtS) '檢查資料正確性

                    If colArray.Length > cst_xc_編號 Then aExamNo3 = colArray(cst_xc_編號).ToString '編號
                    If colArray.Length > cst_xc_科目代號 Then aName = colArray(cst_xc_科目代號).ToString '姓名
                    If colArray.Length > cst_xc_缺考 Then aExam = colArray(cst_xc_缺考).ToString '是否缺考
                    If colArray.Length > cst_xc_成績 Then aWriteResult = colArray(cst_xc_成績).ToString '筆試成績
                    If colArray.Length > cst_xc_口試 Then aOralResult = colArray(cst_xc_口試).ToString '口試成績

                    If Reason = "" Then
                        Dim drS As DataRow = Nothing '取得較多的學員資訊(含編號)
                        Dim dr As DataRow = Nothing
                        'dt3搜尋。
                        aExamNo3 = Right("00" & aExamNo3, 3)
                        If dtS.Select("ExamNo3='" & aExamNo3 & "'").Length > 0 Then
                            drS = dtS.Select("ExamNo3='" & aExamNo3 & "'")(0)
                            'dt要UPDATE
                            If dt.Select("SETID=" & drS("SETID") & " AND EnterDate='" & Common.FormatDate(drS("EnterDate")) & "' AND SerNum=" & drS("SerNum")).Length > 0 Then
                                dr = dt.Select("SETID=" & drS("SETID") & " AND EnterDate='" & Common.FormatDate(drS("EnterDate")) & "' AND SerNum=" & drS("SerNum"))(0)
                            End If
                        End If
                        If dr IsNot Nothing Then
                            Select Case aExam
                                Case "F"
                                    '總成績計算方式
                                    Dim iTotalVal As Double = -1
                                    Dim iWriteResult As Double = 0
                                    Dim iOralResult As Double = 0

                                    Try
                                        'iWriteResult = Val(aWriteResult) '筆試成績
                                        'iOralResult = Val(aOralResult) '筆試成績
                                        iWriteResult = If(blFlag2, Val(aWriteResult), -1) '2018-09-14 fix 免筆試時存-1
                                        iOralResult = If(blFlag3, Val(aOralResult), -1) '2018-09-14 fix 免口試時存-1

                                        '2018-09-14 fix 免筆試/口試時以0分計算總分
                                        iTotalVal = (IIf(iWriteResult = -1, 0, iWriteResult) * Val(ItemVar1.Value) / 100) + (IIf(iOralResult = -1, 0, iOralResult) * Val(ItemVar2.Value) / 100)
                                    Catch ex As Exception
                                    End Try

                                    If blFlag2 Then dr("WriteResult") = iWriteResult 'Val(aWriteResult) '筆試成績
                                    If blFlag3 Then dr("OralResult") = iOralResult 'Val(aOralResult)  '口試成績
                                    'If iWriteResult = -1 OrElse iWriteResult = -1 Then 
                                    If Not ((blFlag2 AndAlso iWriteResult = -1) OrElse (blFlag3 AndAlso iWriteResult = -1)) Then
                                        dr("TotalResult") = iTotalVal
                                    End If
                                    dr("TotalResult") = iTotalVal
                                Case "T"
                                    '缺考成績為-1
                                    dr("WriteResult") = -1
                                    dr("OralResult") = -1
                                    dr("TotalResult") = -1
                                Case Else
                                    '意外成績為零(暫時多寫)
                                    'dr("WriteResult") = 0
                                    'dr("OralResult") = 0
                                    'dr("TotalResult") = 0

                                    '意外成績為-1(暫時多寫)
                                    dr("WriteResult") = -1
                                    dr("OralResult") = -1
                                    dr("TotalResult") = -1
                            End Select
                            dr("ModifyAcct") = sm.UserInfo.UserID
                            dr("ModifyDate") = Now
                            dr("NotExam") = False '是否免試為否
                        End If
                    Else
                        '錯誤資料，填入錯誤資料表
                        drWrong = dtWrong.NewRow
                        dtWrong.Rows.Add(drWrong)
                        drWrong("Index") = RowIndex
                        drWrong("Name") = aName
                        drWrong("IDNO") = aExamNo3
                        drWrong("Reason") = Reason
                    End If
                Next

                DbAccess.UpdateDataTable(dt, da, trans)
                DbAccess.CommitTrans(trans)
                Call TIMS.CloseDbConn(TransConn)

                '判斷匯出資料是否有誤
                Dim explain As String = ""
                explain += "匯入資料共" & dt_xls.Rows.Count & "筆" & vbCrLf
                explain += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
                explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf

                Dim explain2 As String = ""
                explain2 += "匯入資料共" & dt_xls.Rows.Count & "筆\n"
                explain2 += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆\n"
                explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"

                '開始判別欄位存入------------   End
                If TIMS.dtNODATA(dtWrong) Then
                    Common.MessageBox(Me, explain)
                Else
                    Session("MyWrongTable") = dtWrong
                    Dim x_script As String = String.Concat("<script>if(confirm('", explain2, "是否要檢視原因?')){window.open('SD_02_001_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
                    Page.RegisterStartupScript("", x_script)
                End If
            Catch ex As Exception
                DbAccess.RollbackTrans(trans)
                Call TIMS.CloseDbConn(TransConn)
                '，請再重新確認資料狀況!!
                Common.MessageBox(Me, "匯入錯誤，請再重新確認資料狀況!!")
                'Common.MessageBox(Me, ex.ToString)
                Dim strErrmsg As String = String.Concat("ex.Message:", ex.Message, vbCrLf)
                TIMS.WriteTraceLog(Me, ex, strErrmsg)
                Exit Sub
                'Throw ex
            End Try

        End Using

    End Sub

    'UPDATE STUD_ENTERTYPE
    Sub UPDATE_STUD_ENTERTYPE(ByVal htSS As Hashtable)
        Dim WRITERESULT As String = TIMS.GetMyValue2(htSS, "WRITERESULT")
        Dim ORALRESULT As String = TIMS.GetMyValue2(htSS, "ORALRESULT")
        Dim TOTALRESULT As String = TIMS.GetMyValue2(htSS, "TOTALRESULT")

        Dim SETID As String = TIMS.GetMyValue2(htSS, "SETID")
        Dim ENTERDATE As String = TIMS.GetMyValue2(htSS, "ENTERDATE")
        Dim SERNUM As String = TIMS.GetMyValue2(htSS, "SERNUM")

        '2018-09-14 多匯入“加權3%” & "身分別代碼"
        'Dim EXAMPLUS As String = TIMS.GetMyValue2(htSS, "EXAMPLUS")
        'Dim EIdentityID As String = TIMS.GetMyValue2(htSS, "EIDENTITYID")

        ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
        Dim uSql As String = ""
        uSql &= " UPDATE STUD_ENTERTYPE" & vbCrLf
        uSql &= " SET WRITERESULT = @WRITERESULT ,ORALRESULT = @ORALRESULT ,TOTALRESULT = @TOTALRESULT" & vbCrLf
        'uSql += " ,EXAMPLUS = @EXAMPLUS" & vbCrLf 'uSql += " ,EIdentityID = @EIdentityID" & vbCrLf
        uSql &= " ,MODIFYACCT = @MODIFYACCT ,MODIFYDATE = GETDATE()" & vbCrLf
        uSql &= " WHERE  SETID = @SETID AND ENTERDATE = @ENTERDATE AND SERNUM = @SERNUM" & vbCrLf

        '.Parameters.Add("EXAMPLUS", SqlDbType.VarChar).Value = EXAMPLUS
        '.Parameters.Add("EIdentityID", SqlDbType.VarChar).Value = IIf(EIdentityID = "", Convert.DBNull, EIdentityID.PadLeft(2, "0"))
        Dim Parms As New Hashtable From {
            {"WRITERESULT", Val(WRITERESULT)},
            {"ORALRESULT", Val(ORALRESULT)},
            {"TOTALRESULT", Val(TOTALRESULT)},
            {"ModifyAcct", sm.UserInfo.UserID},
            {"SETID", Val(SETID)},
            {"ENTERDATE", TIMS.Cdate2(ENTERDATE)},
            {"SERNUM", Val(SERNUM)}
        }
        DbAccess.ExecuteNonQuery(uSql, objconn, Parms)
    End Sub

    'DataGrid1_ItemDataBound
    Sub SUtl_SetImgSort(ByVal sSort As String, ByVal i As Integer, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs)
        Dim img As New System.Web.UI.WebControls.Image

        Dim v_SORTV1 As String = If(sSort.IndexOf("desc") = -1, "SortUp", "SortDown")
        img.ImageUrl = String.Concat("../../images/", v_SORTV1, ".gif")
        e.Item.Cells(i).Controls.Add(img)
        '顏色調整
        If i <> cst_准考證號 Then e.Item.Cells(cst_准考證號).ForeColor = Color.Black
        If i <> cst_姓名 Then e.Item.Cells(cst_姓名).ForeColor = Color.Black
        If i <> cst_身分證號碼 Then e.Item.Cells(cst_身分證號碼).ForeColor = Color.Black
        If i <> cst_報名日 Then e.Item.Cells(cst_報名日).ForeColor = Color.Black
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim i As Integer = 0
                Dim chkValue As Boolean = False
                Select Case Convert.ToString(ViewState("sort"))
                    Case "ExamNo", "ExamNo desc"
                        i = cst_准考證號
                        chkValue = True
                    Case "Name", "Name desc"
                        i = cst_姓名
                        chkValue = True
                    Case "IDNO", "IDNO desc"
                        i = cst_身分證號碼
                        chkValue = True
                    Case "RelEnterDate", "RelEnterDate desc"
                        i = cst_報名日
                        chkValue = True
                End Select
                If chkValue Then Call SUtl_SetImgSort(Convert.ToString(ViewState("sort")), i, e)

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'Dim exam As DataTable = GetTrain3()
                Dim TextBox1 As TextBox = e.Item.FindControl("TextBox1")
                Dim TextBox2 As TextBox = e.Item.FindControl("TextBox2")
                Dim TextBox3 As TextBox = e.Item.FindControl("TextBox3")
                '如果有勾選筆試則顯示TEXBOX1
                TextBox1.Enabled = If(Convert.ToString(drv("GETTRAIN3")).Contains("2"), True, False)
                '如果有勾選口試則顯示TEXBOX2
                TextBox2.Enabled = If(Convert.ToString(drv("GETTRAIN3")).Contains("3"), True, False)
                TextBox1.MaxLength = 10
                TextBox2.MaxLength = 10
                TextBox3.MaxLength = 10
                If Not TextBox1.Enabled Then TextBox1.ToolTip = cst_免筆試
                If Not TextBox2.Enabled Then TextBox1.ToolTip = cst_免口試

                'If Exam.Rows(0)("GetTrain3").ToString().Contains("2") Then TextBox1.Enabled = True '如果有勾選筆試則顯示TEXBOX1
                'If Exam.Rows(0)("GetTrain3").ToString().Contains("3") Then TextBox2.Enabled = True '如果有勾選口試則顯示TEXBOX2
                'Dim NotExam As HtmlInputCheckBox = e.Item.FindControl("NotExam")

                '甄試加分(加權3%) EXAMPLUS
                'Dim EXAMPLUS As HtmlInputCheckBox = e.Item.FindControl("EXAMPLUS")
                '身分別
                'Dim TYPE_EIdentityID As DropDownList = e.Item.FindControl("TYPE_EIdentityID")
                'Dim Hid_IdentityID As HiddenField = e.Item.FindControl("Hid_IdentityID")

                'Dim L_TRNDType As Label = e.Item.FindControl("L_TRNDType")
                Dim Hid_SETID As HtmlInputHidden = e.Item.FindControl("Hid_SETID")
                Dim Hid_EnterDate As HtmlInputHidden = e.Item.FindControl("Hid_EnterDate")
                Dim Hid_SerNum As HtmlInputHidden = e.Item.FindControl("Hid_SerNum")
                Hid_SETID.Value = Convert.ToString(drv("SETID"))
                Hid_EnterDate.Value = TIMS.Cdate3(Convert.ToString(drv("EnterDate")))
                Hid_SerNum.Value = Convert.ToString(drv("SerNum"))

                'TIMS.Tooltip(EXAMPLUS, "甄試加分*1.03")
                'NotExam.Attributes("onclick") = "NotExam(this.checked," & e.Item.ItemIndex + 1 & ")"
                'Dim blnNotExam As Boolean = False '免試 booleam
                'blnNotExam = drv("NotExam") '免試
                'EXAMPLUS.Checked = False '甄試加分
                'If Convert.ToString(drv("EXAMPLUS")) = "1" Then EXAMPLUS.Checked = True '甄試加分
                'TYPE_EIdentityID = TIMS.Get_Identity(TYPE_EIdentityID, 5, sm.UserInfo.TPlanID, sm.UserInfo.Years, objconn)
                'Hid_IdentityID.Value = ""
                'If Convert.ToString(drv("EIdentityID")) <> "" Then
                '    Common.SetListItem(TYPE_EIdentityID, drv("EIdentityID"))
                '    'Hid_IdentityID.Value = TYPE_IdentityID.SelectedValue
                'End If

                '券別
                'Dim TRNDTypeVal As String = "-"
                'Select Case Convert.ToString(drv("TRNDType"))
                '    Case "1"
                '        TRNDTypeVal = "甲式"
                '    Case "2"
                '        TRNDTypeVal = "乙式"
                'End Select
                'If Convert.ToString(drv("WSort")) = "1" Then
                '    'blnNotExam = True
                '    'blnNotExam = False '本來免輸入2016要輸入成績
                '    TRNDTypeVal = "就服單位協助報名"
                'End If
                'e.Item.Cells(cst_券別).Text = TRNDTypeVal
                'L_TRNDType.Text = TRNDTypeVal

                'NotExam.Checked = blnNotExam '免試
                'TextBox1.Enabled = Not blnNotExam '免輸入
                'TextBox2.Enabled = Not blnNotExam '免輸入
                'TextBox3.Enabled = Not blnNotExam '免輸入

                'If Convert.ToString(drv("GOVKILL")) = "Y" Then
                '    e.Item.Enabled = False
                '    TIMS.Tooltip(e.Item, cst_Mgc219)
                'End If
        End Select
    End Sub

    Private Sub DataGrid1_SortCommand(source As Object, e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid1.SortCommand
        ViewState("sort") = String.Concat(e.SortExpression, If(ViewState("sort") = e.SortExpression, " desc", ""))

        'Button1_Click(Button1, Nothing)
        Call Search1()
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        'Dim obj As DataGrid = sender
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim stud As TextBox = e.Item.FindControl("stud1")
                stud.Text = TIMS.Get_DGSeqNo(sender, e) '序號
        End Select
    End Sub

    'Button2.Attributes("onclick") 
    'Protected Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    'End Sub

    Protected Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged
    End Sub

    ''' <summary>
    '''  匯出身分別代碼
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnExportIdentity_Click(sender As Object, e As EventArgs) Handles btnExportIdentity.Click
        Dim dt As DataTable = GetIdentityID()
        ExportIdentityID(dt)
    End Sub

    ''' <summary>
    ''' 匯出身分別代碼表
    ''' </summary>
    ''' <param name="dt"></param>
    Sub ExportIdentityID(ByVal dt As DataTable)
        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode("身分別代碼對照表", System.Text.Encoding.UTF8) & ".xls")
        Response.ContentType = "Application/octet-stream"
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")

        Dim exportStr As String = ""            '建立輸出文字
        exportStr &= String.Concat("身分別代碼", vbTab, "身分別名稱", vbTab, vbCrLf)
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(exportStr))

        For Each dr As DataRow In dt.Rows
            Dim exportStr2 As String = String.Concat(dr("IDENTITYID"), vbTab, dr("NAME"), vbTab, vbCrLf)
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(exportStr2))
        Next

        TIMS.CloseDbConn(objconn)
        Response.End()
    End Sub

End Class