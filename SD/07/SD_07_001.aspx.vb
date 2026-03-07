Partial Class SD_07_001
    Inherits AuthBasePage

    'STUD_TECHEXAM 'KEY_EXAM 'SELECT * FROM V_STUDTECHEXAM WHERE OCID='57518' 
    'KEY_EXAM/CLASS_TECHEXAM/STUD_TECHEXAM3
    'V_STUDEXAM1 / V_STUDEXAM2 / V_STUDEXAM3

    Const cst_ss_Sd07001_sch1 As String = "ss_Sd07001_sch1"
    Const cst_申請設定 As String = "setup1"
    Const cst_結果輸入 As String = "setup2"

    Dim dtStdTechExam As DataTable

    Dim vMsg As String = ""

    Const cst_xExamLevelv1 As String = "1,2,3,4,5"
    Const cst_xExamLeveln1 As String = "甲級,乙級,丙級,單一級,不分級"

    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

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
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        '檢查Session是否存在 End

        Call CreateJScript()

        If Not IsPostBack Then
            msg.Text = ""
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            btnSaveData1.Visible = False '儲存
            btnExport1.Visible = False '匯出
            btnBack1.Visible = False '回上頁
        End If

        '檢查帳號的功能權限-----------------------------------Start
        'btnAdd1.Enabled = True '新增
        'If Not blnCanAdds Then
        '    btnAdd1.Enabled = False
        '    TIMS.Tooltip(btnAdd1, "(Adds)無權限使用該功能", True)
        'End If
        'btnSearch.Enabled = True '查詢
        'If Not blnCanSech Then
        '    btnSearch.Enabled = False
        '    TIMS.Tooltip(btnSearch, "(Sech) 無權限使用該功能", True)
        'End If
        '檢查帳號的功能權限-----------------------------------End

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        btnSetLevOrg.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True)
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        btnSearch.Attributes("onclick") = "javascript:return search();"
        btnAdd1.Attributes("onclick") = "javascript:return addnew1();"
        tExamKind.Attributes("onblur") = "Get_Exam(this,'" & tExamName.ClientID & "');"
    End Sub

    '查詢
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Call keepSearch1()
        ' Common.MessageBox(Me, "查詢")
        Call reloadSch1()
        '查詢
        Call show_dg1(OCIDValue1.Value, tExamKind.Text)
    End Sub

    '申請設定
    Private Sub btnAdd1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd1.Click
        Call reloadSch1() '1.重新查詢

        Call show_dg2()
    End Sub

    '刪除動作1
    Sub sDeleteData1(ByVal sCmdArg As String)
        'Dim sCmdArg As String = e.CommandArgument
        Dim CTEID As String = TIMS.GetMyValue(sCmdArg, "CTEID")
        Dim OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        If CTEID = "" Then Exit Sub
        If OCID = "" Then Exit Sub

        Call TIMS.OpenDbConn(objconn)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select 'X'" & vbCrLf
        sql &= " FROM CLASS_TECHEXAM ct" & vbCrLf
        sql &= " JOIN STUD_TECHEXAM3 c on c.CTEID=ct.CTEID" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs on cs.socid =c.socid" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss on ss.sid =cs.sid" & vbCrLf
        sql &= " JOIN KEY_EXAM k1 on k1.EXAMID=CT.EXAMKIND" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and c.PASS IS NOT NULL" & vbCrLf
        sql &= " and ct.OCID=@OCID" & vbCrLf
        sql &= " and ct.CTEID=@CTEID" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID
            .Parameters.Add("CTEID", SqlDbType.VarChar).Value = CTEID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            Common.MessageBox(Me, "尚有學員技檢資料不可刪除!!")
            Exit Sub
        End If

        sql = "" & vbCrLf
        sql &= " DELETE STUD_TECHEXAM3" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " and CTEID=@CTEID" & vbCrLf
        Dim dCmd1 As New SqlCommand(sql, objconn)
        With dCmd1
            .Parameters.Clear()
            .Parameters.Add("CTEID", SqlDbType.VarChar).Value = CTEID
            '.ExecuteNonQuery()
            DbAccess.ExecuteNonQuery(dCmd1.CommandText, objconn, dCmd1.Parameters)
        End With

        sql = "" & vbCrLf
        sql &= " DELETE CLASS_TECHEXAM" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " and OCID=@OCID" & vbCrLf
        sql &= " and CTEID=@CTEID" & vbCrLf
        Dim dCmd2 As New SqlCommand(sql, objconn)
        With dCmd2
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID
            .Parameters.Add("CTEID", SqlDbType.VarChar).Value = CTEID
            '.ExecuteNonQuery()
            DbAccess.ExecuteNonQuery(dCmd2.CommandText, objconn, dCmd2.Parameters)
        End With
    End Sub

    '1.結果輸入
    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub

        Call keepSearch1()
        Select Case e.CommandName
            Case "btnEdit2"
                Call reloadSch1()
                '結果輸入
                Call show_dg3(e.CommandArgument)

            Case "btnDel2"
                Dim sCmdArg As String = e.CommandArgument
                '刪除動作1
                Call sDeleteData1(sCmdArg)

                Call reloadSch1()
                '查詢
                Call show_dg1(OCIDValue1.Value, tExamKind.Text)

        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btnEdit2 As Button = e.Item.FindControl("btnEdit2")
                Dim btnDel2 As Button = e.Item.FindControl("btnDel2")
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號 

                btnDel2.Enabled = False
                If Val(drv("CNT1")) = 0 Then btnDel2.Enabled = True
                If Not btnDel2.Enabled Then TIMS.Tooltip(btnDel2, "尚有資料不可刪除!!")

                Dim sCmdArg As String = ""
                sCmdArg = "" '結果輸入
                TIMS.SetMyValue(sCmdArg, "CTEID", drv("CTEID"))
                TIMS.SetMyValue(sCmdArg, "OCID", drv("OCID"))
                TIMS.SetMyValue(sCmdArg, "ExamKind", drv("ExamKind"))
                TIMS.SetMyValue(sCmdArg, "ExamTime", drv("ExamTime"))
                btnEdit2.CommandArgument = sCmdArg

                If btnDel2.Enabled Then
                    btnDel2.CommandArgument = sCmdArg
                    TIMS.Tooltip(btnDel2, "未輸入詳細資料，提供刪除功能。")
                End If
                If Not btnDel2.Enabled Then btnDel2.CommandArgument = "" '不提供刪除功能

        End Select
    End Sub

    '查詢
    Sub show_dg1(ByVal ocid As String, ByVal examkind As String)
        '技能檢定的基礎設定

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " select ct.ocid,ct.cteid" & vbCrLf
        sql &= " ,count(c.pass) cnt1" & vbCrLf
        sql &= " FROM CLASS_TECHEXAM ct" & vbCrLf
        sql &= " JOIN STUD_TECHEXAM3 c on c.CTEID=ct.CTEID" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs on cs.socid =c.socid" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss on ss.sid =cs.sid" & vbCrLf
        sql &= " JOIN KEY_EXAM k1 on k1.EXAMID=CT.EXAMKIND" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        If ocid <> "" Then
            sql &= " and ct.ocid='" & ocid & "'" & vbCrLf
        End If
        sql &= " group by ct.ocid,ct.cteid" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " select cc.classcname" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111) stdate" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.ftdate, 111) ftdate" & vbCrLf
        sql &= " ,cc.ocid" & vbCrLf
        sql &= " ,ct.CTEID" & vbCrLf
        sql &= " ,ct.ExamKind" & vbCrLf
        sql &= " ,ct.ExamTime" & vbCrLf
        sql &= " ,k1.name ExamName" & vbCrLf
        sql &= " ,IsNull(wc1.cnt1,0) cnt1" & vbCrLf

        sql &= " FROM CLASS_CLASSINFO cc" & vbCrLf
        '依 ocid (或 +ExamKind )
        sql &= " JOIN CLASS_TECHEXAM ct on ct.ocid =cc.ocid" & vbCrLf
        sql &= " JOIN KEY_EXAM k1 on k1.EXAMID=ct.ExamKind" & vbCrLf
        sql &= " LEFT JOIN WC1 on wc1.CTEID=ct.CTEID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If ocid <> "" Then
            sql &= " and cc.OCID='" & ocid & "'" & vbCrLf
        End If
        If examkind <> "" Then
            sql &= " and ct.ExamKind ='" & examkind & "'"
        End If
        sql &= " order by ct.ExamTime" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        DataGrid1c.Visible = False
        DataGrid1b.Visible = False
        DataGrid1.Visible = False

        msg3.Text = ""
        msg2.Text = ""
        msg.Text = "查無資料!!"
        If dt.Rows.Count > 0 Then
            DataGrid1.Visible = True
            msg.Text = ""

            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If
    End Sub

    '申請設定
    Sub show_dg2()
        Dim iExamTime As String = 1 '可以輸入 1, 2, 3 (目前系統限定3種職類，以班為單位)
        hid_CTEID.Value = ""

        '驗證可能產生的錯誤。
        Dim errMsg As String = ""
        Dim ocid As String = Me.OCIDValue1.Value
        Dim ExamKind As String = Me.tExamKind.Text
        ExamKind = TIMS.ClearSQM(ExamKind)
        If ocid = "" Then errMsg += "請選擇班級職類" & vbCrLf
        If ExamKind = "" Then errMsg += "[新增]請輸入選擇檢定職類" & vbCrLf
        'hid_CTEID.Value = ""
        hid_OCID.Value = ocid
        hid_EXAMKIND.Value = ExamKind

        Me.hid_ExamTime.Value = ""
        If errMsg = "" Then
            '取得為第幾次檢定資料 
            iExamTime = sUtl_GetExamTime(ocid, ExamKind)
            'Select Case iExamTime
            '    Case 4
            '        errMsg += "[新增]已超過系統上限檢定資料次數(3)" & vbCrLf
            'End Select
        End If

        'ExamTime
        If errMsg <> "" Then
            Common.MessageBox(Me, errMsg)
            Exit Sub
        End If
        Me.hid_ExamTime.Value = iExamTime

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.StudentID" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(a.StudentID) STUDID2" & vbCrLf '學號
        sql &= " ,a.SOCID" & vbCrLf
        sql &= " ,a.OCID" & vbCrLf
        sql &= " ,a.StudStatus" & vbCrLf
        sql &= " ,b.Name" & vbCrLf
        sql &= " ,ct.CTEID" & vbCrLf
        sql &= " ,ct.ExamKind" & vbCrLf
        sql &= " ,k1.name ExamName" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS a" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO b on a.SID=b.SID" & vbCrLf
        '依 ocid,ExamKind,ExamTime
        sql &= " LEFT JOIN CLASS_TECHEXAM ct ON ct.OCID =a.OCID" & vbCrLf
        sql &= " and ct.ExamKind ='" & ExamKind & "'" & vbCrLf
        sql &= " and ct.ExamTime ='" & iExamTime & "'" & vbCrLf
        sql &= " LEFT JOIN KEY_EXAM k1 on k1.EXAMID=ct.ExamKind" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " and a.OCID =" & ocid & vbCrLf
        sql &= " ORDER BY a.StudentID" & vbCrLf
        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn)

        dtStdTechExam = sUtl_GetStdTechExam(ocid, ExamKind, iExamTime)

        tExamKind.Enabled = True

        btnSaveData1.CommandName = ""
        btnExport1.CommandName = ""
        btnBack1.CommandName = ""

        btnSaveData1.Visible = False '儲存
        btnExport1.Visible = False '匯出
        btnBack1.Visible = False '回上頁

        DataGrid1c.Visible = False
        DataGrid1.Visible = False
        DataGrid1b.Visible = False

        msg.Text = ""
        msg3.Text = ""
        msg2.Text = "查無資料!!"
        If dt.Rows.Count > 0 Then
            hid_CTEID.Value = Convert.ToString(dt.Rows(0)("CTEID"))
            tExamKind.Enabled = False

            btnSaveData1.CommandName = cst_申請設定 '儲存
            btnExport1.CommandName = cst_申請設定 '匯出
            btnBack1.CommandName = cst_申請設定 '回上頁

            btnSaveData1.Visible = True '儲存
            btnExport1.Visible = True '匯出
            btnBack1.Visible = True '回上頁

            DataGrid1b.Visible = True
            msg2.Text = ""

            DataGrid1b.DataSource = dt
            DataGrid1b.DataBind()
        End If

    End Sub

    '結果輸入
    Sub show_dg3(ByVal cmdArg As String)
        Dim errMsg As String = ""
        Dim ocid As String = TIMS.GetMyValue(cmdArg, "ocid")
        Dim CTEID As String = TIMS.GetMyValue(cmdArg, "CTEID")
        Dim ExamKind As String = TIMS.GetMyValue(cmdArg, "ExamKind")
        Dim ExamTime As String = TIMS.GetMyValue(cmdArg, "ExamTime")
        If ocid = "" Then Exit Sub
        If CTEID = "" Then Exit Sub
        If ExamKind = "" Then Exit Sub
        If ExamTime = "" Then Exit Sub
        Dim flagChk123 As Boolean = TIMS.Check123(ExamTime)
        If Not flagChk123 Then errMsg &= "輸入資料有誤!!" & vbCrLf
        'Select Case ExamTime
        '    Case "1", "2", "3"
        '    Case Else
        '        errMsg += "輸入資料有誤!!" & vbCrLf
        'End Select
        If errMsg <> "" Then
            Common.MessageBox(Me, errMsg)
            Exit Sub
        End If

        hid_EXAMKIND.Value = ExamKind
        hid_CTEID.Value = CTEID
        hid_ExamTime.Value = ExamTime
        hid_OCID.Value = ocid

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT c.STXID" & vbCrLf
        sql &= " ,c.CTEID" & vbCrLf
        sql &= " ,c.SOCID" & vbCrLf
        sql &= " ,c.EXAMLEVEL" & vbCrLf
        sql &= " ,CONVERT(varchar, c.APPLYDATE, 111) APPLYDATE" & vbCrLf
        sql &= " ,c.PASS" & vbCrLf
        sql &= " ,CONVERT(varchar, c.SENDOUTCERTDATE, 111) SENDOUTCERTDATE" & vbCrLf
        sql &= " ,c.EXAMNO" & vbCrLf
        sql &= " ,CONVERT(varchar, c.EXAMDATE, 111) EXAMDATE" & vbCrLf
        sql &= " ,ct.EXAMKIND" & vbCrLf
        sql &= " ,k1.name EXAMNAME" & vbCrLf
        sql &= " ,ss.Name" & vbCrLf
        sql &= " ,cs.StudentID" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(cs.StudentID) STUDID2" & vbCrLf
        sql &= " ,cs.StudStatus" & vbCrLf
        '依 ocid,CTEID
        sql &= " FROM CLASS_TECHEXAM ct" & vbCrLf
        sql &= " JOIN STUD_TECHEXAM3 c on c.CTEID=ct.CTEID" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs on cs.socid =c.socid" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss on ss.sid =cs.sid" & vbCrLf
        sql &= " JOIN KEY_EXAM k1 on k1.EXAMID=CT.EXAMKIND" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " and ct.OCID =" & ocid & vbCrLf
        sql &= " and ct.CTEID =" & CTEID & vbCrLf
        sql &= " ORDER BY cs.StudentID" & vbCrLf
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        tExamKind.Enabled = True

        btnSaveData1.CommandName = ""
        btnExport1.CommandName = ""
        btnBack1.CommandName = ""

        btnSaveData1.Visible = False '儲存
        btnExport1.Visible = False '匯出
        btnBack1.Visible = False '回上頁

        DataGrid1.Visible = False
        DataGrid1b.Visible = False
        DataGrid1c.Visible = False

        msg.Text = ""
        msg2.Text = ""
        msg3.Text = "查無資料!!"
        If dt.Rows.Count > 0 Then
            tExamKind.Enabled = False
            tExamKind.Text = dt.Rows(0)("ExamKind")
            tExamName.Text = dt.Rows(0)("ExamName")
            hid_ExamTime.Value = ExamTime

            btnSaveData1.CommandName = cst_結果輸入
            btnExport1.CommandName = cst_結果輸入
            btnBack1.CommandName = cst_結果輸入

            btnSaveData1.Visible = True '儲存
            btnExport1.Visible = True '匯出
            btnBack1.Visible = True '回上頁

            DataGrid1c.Visible = True
            msg3.Text = ""

            DataGrid1c.DataSource = dt
            DataGrid1c.DataBind()
        End If
    End Sub

    '申請設定
    Private Sub DataGrid1b_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1b.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim myCheckbox2 As HtmlInputCheckBox = e.Item.FindControl("Checkbox2") '全選 HtmlInputCheckBox
                Dim cbl1TrainLevel As CheckBoxList = e.Item.FindControl("cbl1TrainLevel") '檢定級別
                Dim hidSOCID As HtmlInputHidden = e.Item.FindControl("hidSOCID")
                hidSOCID.Value = Convert.ToString(drv("SOCID"))

                Dim sEXAMLEVELS As String = Get_Examlevels1(dtStdTechExam, hidSOCID.Value, hid_ExamTime.Value)
                TIMS.SetCblValue(cbl1TrainLevel, sEXAMLEVELS)

                '申請日  tApplyDate(IMG1)
                Dim tApplyDate As TextBox = e.Item.FindControl("tApplyDate")
                Dim sApplyDate As String = Get_ApplyDate1(dtStdTechExam, hidSOCID.Value, hid_ExamTime.Value)
                tApplyDate.Text = TIMS.Cdate3(sApplyDate)
                Dim myIMG1 As HtmlImage = e.Item.FindControl("IMG1")
                myIMG1.Attributes("onclick") = "javascript:show_calendar('" & tApplyDate.ClientID & "','','','CY/MM/DD');"

                myCheckbox2.Disabled = False
                'myCheckbox3.Disabled = False
                'myDropDownList1.Enabled = True
                cbl1TrainLevel.Enabled = True '檢定級別
                'myIMG1.Style.Item("display") = "inline"
                Call Get_ExamPass2(dtStdTechExam, hidSOCID.Value, hid_ExamTime.Value, cbl1TrainLevel)

                'Select Case sExamPass1 'Convert.ToString(drv("Pass"))
                '    Case "Y" '檢定結果-通過
                '        myCheckbox2.Disabled = True
                '        TIMS.Tooltip(myCheckbox2, "此學員 檢定結果-通過", True)
                '        tApplyDate.Enabled = False '申請日 
                '        myIMG1.Style.Item("display") = "none"
                '        TIMS.Tooltip(tApplyDate, "此學員 檢定結果-通過", True)
                '        'myDropDownList1.Enabled = False
                '        cbl1TrainLevel.Enabled = False
                '        TIMS.Tooltip(cbl1TrainLevel, "此學員 檢定結果-通過", True)
                '    Case "N" 'N:不通過
                '        myCheckbox2.Disabled = True
                '        TIMS.Tooltip(myCheckbox2, "此學員 檢定結果-不通過", True)
                '        tApplyDate.Enabled = False '申請日 
                '        myIMG1.Style.Item("display") = "none"
                '        TIMS.Tooltip(tApplyDate, "此學員 檢定結果-不通過", True)
                '        cbl1TrainLevel.Enabled = False
                '        TIMS.Tooltip(cbl1TrainLevel, "此學員 檢定結果-不通過", True)
                '    Case Else '檢定結果-NULL:缺考 
                '        TIMS.Tooltip(myCheckbox2, "此學員 檢定結果-缺", True)
                '        TIMS.Tooltip(tApplyDate, "此學員 檢定結果-缺", True)
                'End Select

                '學員已離退訓
                Select Case Convert.ToString(drv("StudStatus"))
                    Case "2", "3"
                        'myDropDownList1.Enabled = False '檢定級別
                        myCheckbox2.Disabled = True
                        TIMS.Tooltip(myCheckbox2, "此學員已離退訓", True)

                        tApplyDate.Enabled = False '申請日 
                        TIMS.Tooltip(tApplyDate, "此學員已離退訓", True)
                        myIMG1.Style.Item("display") = "none"
                        'myDropDownList1.Enabled = False
                        cbl1TrainLevel.Enabled = False '檢定級別
                        TIMS.Tooltip(cbl1TrainLevel, "此學員已離退訓", True)

                End Select

                myCheckbox2.Attributes("onclick") = "select_all(3,this.checked," & e.Item.ItemIndex + 2 & ");"
        End Select

    End Sub

    '結果輸入
    Private Sub DataGrid1c_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1c.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem

                '級別
                Dim lExamLevel As Label = e.Item.FindControl("lExamLevel")
                Dim slevel As String = ""
                If Convert.ToString(drv("EXAMLEVEL")) <> "" Then
                    slevel = sGet_ExamLevels(drv("EXAMLEVEL"))
                End If
                lExamLevel.Text = slevel
                'e.Item.Cells(3).Text = slevel

                '檢定結果
                Dim ddlExamPass As DropDownList = e.Item.FindControl("ddlExamPass")
                If Convert.ToString(drv("Pass")) <> "" Then
                    Common.SetListItem(ddlExamPass, drv("Pass"))
                End If

                '檢定日 Textbox6 (Img3)
                '製證日 TextBox4 (IMG2)
                '證號 TextBox5
                Dim Textbox6 As TextBox = e.Item.FindControl("Textbox6")
                Dim TextBox4 As TextBox = e.Item.FindControl("TextBox4")
                Dim TextBox5 As TextBox = e.Item.FindControl("TextBox5")

                Dim myImg3 As HtmlImage = e.Item.FindControl("Img3")
                myImg3.Attributes("onclick") = "javascript:show_calendar('" & Textbox6.ClientID & "','','','CY/MM/DD');"
                'Myimg1.Attributes("onclick") = "javascript:show_calendar('DataGrid2__ctl" & e.Item.ItemIndex + 2 & "_Textbox6','','','CY/MM/DD');"

                Dim myIMG2 As HtmlImage = e.Item.FindControl("IMG2")
                myIMG2.Attributes("onclick") = "javascript:show_calendar('" & TextBox4.ClientID & "','','','CY/MM/DD');"
                'Myimg.Attributes("onclick") = "javascript:show_calendar('DataGrid2__ctl" & e.Item.ItemIndex + 2 & "_TextBox4','','','CY/MM/DD');"

        End Select
    End Sub

    '匯出
    Sub sUtl_Export1(ByVal ocid_Val As String, ByVal examkindval As String, ByVal examtime As String)
        'Const Cst_ExcelSample As String = "sample2.xls"

        'sql &= " ,DATEPART(YEAR, CONVERT(numeric, to_char(b.Birthday)) - 1911,'099')+to_char(b.Birthday,'MMDD') Birthday" & vbCrLf
        'sql &= " ,CONVERT(varchar, c.zipcode1)+case when c.ZipCode1_6W <10 then '0'+CONVERT(varchar, c.ZipCode1_6W) else CONVERT(varchar, c.ZipCode1_6W) end ZipCode1" & vbCrLf

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.SOCID" & vbCrLf
        sql &= " ,a.OCID" & vbCrLf
        sql &= " ,a.StudentID" & vbCrLf
        sql &= " ,dbo.fn_CSTUDID2(a.StudentID) STUDID2" & vbCrLf
        sql &= " ,RIGHT('000'+isnull(convert(varchar,(DATEPART(YEAR,st.ApplyDate)-1911)),'099'),3) ApplyYear" & vbCrLf
        sql &= " ,ct.ExamKind" & vbCrLf
        sql &= " ,st.ExamLevel" & vbCrLf
        sql &= " ,b.IDNO" & vbCrLf
        sql &= " ,b.Name" & vbCrLf
        sql &= " ,b.Engname" & vbCrLf
        sql &= " ,case when b.Sex='M' then '1' else '2' end Sex" & vbCrLf
        sql &= " ,b.Birthday" & vbCrLf
        sql &= " ,dbo.RT_DataFormat(b.Birthday) Birthday" & vbCrLf
        sql &= " ,CASE (b.DegreeID)" & vbCrLf
        sql &= " When '01' Then '21'" & vbCrLf
        sql &= " When '02' Then '32'" & vbCrLf
        sql &= " When '03' Then '41'" & vbCrLf
        sql &= " When '04' Then '50'" & vbCrLf
        sql &= " When '05' Then '60'" & vbCrLf
        sql &= " When '06' Then '70' Else '90' End DegreeID" & vbCrLf
        sql &= " ,b.GraduateStatus" & vbCrLf
        sql &= " ,ISNULL(c.PhoneD,' ') PhoneD" & vbCrLf
        sql &= " ,ISNULL(c.PhoneN,' ') PhoneN" & vbCrLf
        sql &= " ,ISNULL(c.CellPhone,' ') CellPhone" & vbCrLf
        sql &= " ,c.Email" & vbCrLf
        sql &= " ,dbo.FN_GET_ZIPCODE(c.zipcode1,c.ZipCode1_6W) ZipCode1" & vbCrLf
        sql &= " ,iz.CTName CTName1" & vbCrLf
        sql &= " ,iz.ZName ZipName1" & vbCrLf
        sql &= " ,c.Address Address1" & vbCrLf
        sql &= " ,dbo.FN_GET_ZIPCODE(c.ZipCode2,c.ZipCode2_6W) ZipCode2" & vbCrLf
        sql &= " ,iz2.CTName CTName2" & vbCrLf
        sql &= " ,iz2.ZName ZipName2" & vbCrLf
        sql &= " ,c.HouseholdAddress Address2" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS a" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO b on a.SID=b.SID" & vbCrLf
        sql &= " JOIN STUD_SUBDATA c on a.SID=c.SID" & vbCrLf
        sql &= " JOIN STUD_TECHEXAM3 st on st.SOCID =a.SOCID" & vbCrLf
        sql &= " JOIN CLASS_TECHEXAM ct on ct.CTEID=st.CTEID and ct.OCID=a.OCID" & vbCrLf
        sql &= " LEFT JOIN VIEW_ZIPNAME iz on iz.ZipCode = c.ZipCode1" & vbCrLf
        sql &= " LEFT JOIN VIEW_ZIPNAME iz2 on iz2.ZipCode = c.ZipCode2" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf

        sql &= " and a.OCID='" & ocid_Val & "'" & vbCrLf
        sql &= " and ct.EXAMTIME=" & examtime & vbCrLf
        sql &= " ORDER BY a.StudentID" & vbCrLf

        'OCID =73371 
        '(SELECT * FROM Class_StudentsOfClass WHERE OCID ='" & ocidvalue1 & "'" & vbCrLf
        ''sql &= " --WHERE OCID='23349'" & vbCrLf
        'sql &= " ) a " & vbCrLf
        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料可匯出")
            Exit Sub
        End If
        Const cst_SampleXLS As String = "~\SD\07\sample2.xls" ' & Cst_ExcelSample
        If Not IO.File.Exists(Server.MapPath(cst_SampleXLS)) Then
            Common.MessageBox(Me, "sample檔案不存在!")
            Exit Sub
        End If

        Dim strErrmsg As String = ""
        Dim sFileName As String = String.Concat("~\SD\07\", TIMS.GetDateNo(), ".xls")
        Dim MyPath As String = Server.MapPath(sFileName)
        Try
            IO.File.Copy(Server.MapPath(cst_SampleXLS), MyPath, True)
            IO.File.SetAttributes(MyPath, IO.FileAttributes.Normal)
        Catch ex As Exception
            strErrmsg = ""
            strErrmsg += "目錄名稱或磁碟區標籤語法錯誤!!!" & vbCrLf
            strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉) " & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            Common.MessageBox(Me, strErrmsg)
            'Exit Sub
        End Try

        'Dim MyFileName As String = Replace(Replace(Replace(OCID1.Text, ")", ""), "(", ""), "/", "") & ".xls"
        'Dim MyPath As String = Server.MapPath("~\SD\07\" & MyFileName)
        'sFileName = ""
        'sFileName += "~\SD\07\"
        ''sFileName += TIMS.ChangeIDNO(Replace(Replace(Replace(OCID1.Text, ")", ""), "(", ""), "/", ""))
        'sFileName += TIMS.GetDateNo()
        'sFileName += ".xls"   Dim MyPath As String = ""

        Dim MyFileName As String = TIMS.ChangeIDNO(Replace(Replace(Replace(OCID1.Text, ")", ""), "(", ""), "/", "")) & ".xls"
        Using MyConn As New OleDb.OleDbConnection
            MyConn.ConnectionString = TIMS.Get_OleDbStr(MyPath)
            Try
                MyConn.Open()
            Catch
                Common.MessageBox(Me, "Excel資料無法開啟連線!")
                Exit Sub
            End Try

            For Each drv As DataRow In dt.Rows
                Dim STUDENTID As String = ReplaceText(Convert.ToString(drv("STUDENTID")))
                Dim APPLYYEAR As String = ReplaceText(Convert.ToString(drv("APPLYYEAR")))
                Dim EXAMKIND As String = ReplaceText(Convert.ToString(drv("EXAMKIND")))
                Dim EXAMLEVEL As String = ReplaceText(Convert.ToString(drv("EXAMLEVEL")))
                Dim IDNO As String = ReplaceText(Convert.ToString(drv("IDNO")))
                Dim NAME As String = ReplaceText(Convert.ToString(drv("NAME")))
                Dim ENGNAME As String = ReplaceText(Convert.ToString(drv("ENGNAME")))
                Dim SEX As String = ReplaceText(Convert.ToString(drv("SEX")))
                Dim BIRTHDAY As String = ReplaceText(Convert.ToString(drv("BIRTHDAY")))
                Dim DEGREEID As String = ReplaceText(Convert.ToString(drv("DEGREEID")))
                Dim GRADUATESTATUS As String = ReplaceText(Convert.ToString(drv("GRADUATESTATUS")))
                Dim PHONED As String = ReplaceText(Convert.ToString(drv("PHONED")))
                Dim PHONEN As String = ReplaceText(Convert.ToString(drv("PHONEN")))
                Dim CELLPHONE As String = ReplaceText(Convert.ToString(drv("CELLPHONE")))
                Dim EMAIL As String = ReplaceText(Convert.ToString(drv("EMAIL")))
                Dim ZIPCODE1 As String = ReplaceText(Convert.ToString(drv("ZIPCODE1")))
                Dim CTNAME1 As String = ReplaceText(Convert.ToString(drv("CTNAME1")))
                Dim ZIPNAME1 As String = ReplaceText(Convert.ToString(drv("ZIPNAME1")))
                Dim ADDRESS1 As String = ReplaceText(Convert.ToString(drv("ADDRESS1")))
                Dim ZIPCODE2 As String = ReplaceText(Convert.ToString(drv("ZIPCODE2")))
                Dim CTNAME2 As String = ReplaceText(Convert.ToString(drv("CTNAME2")))
                Dim ZIPNAME2 As String = ReplaceText(Convert.ToString(drv("ZIPNAME2")))
                Dim ADDRESS2 As String = ReplaceText(Convert.ToString(drv("ADDRESS2")))

                If Not ADDRESS1.IndexOf(CTNAME1 & ZIPNAME1) > -1 Then
                    ADDRESS1 = CTNAME1 & ZIPNAME1 & ADDRESS1
                End If
                If Not ADDRESS2.IndexOf(CTNAME2 & ZIPNAME2) > -1 Then
                    ADDRESS2 = CTNAME2 & ZIPNAME2 & ADDRESS2
                End If

                sql = "Insert Into [Sheet1$] ("
                sql &= " EXAMKIND" & vbCrLf
                sql &= " ,SOPCD" & vbCrLf
                sql &= " ,PID" & vbCrLf
                sql &= " ,YR" & vbCrLf
                sql &= " ,STP" & vbCrLf
                sql &= " ,PNO" & vbCrLf
                sql &= " ,EGR" & vbCrLf
                sql &= " ,AENO" & vbCrLf
                sql &= " ,DSTNG" & vbCrLf
                sql &= " ,ADSTNG" & vbCrLf
                sql &= " ,BDSTNG1" & vbCrLf
                sql &= " ,IDNO" & vbCrLf
                sql &= " ,[NAME]" & vbCrLf
                sql &= " ,ENNAME" & vbCrLf
                sql &= " ,SEX" & vbCrLf
                sql &= " ,BIRDTE" & vbCrLf
                sql &= " ,WEAKID" & vbCrLf
                sql &= " ,EDU" & vbCrLf
                sql &= " ,EDUSTS" & vbCrLf
                sql &= " ,OTEL" & vbCrLf
                sql &= " ,HTEL" & vbCrLf
                sql &= " ,MTEL" & vbCrLf
                sql &= " ,EMAIL" & vbCrLf
                sql &= " ,CZIP" & vbCrLf
                sql &= " ,CADR" & vbCrLf
                sql &= " ,FZIP" & vbCrLf
                sql &= " ,FADR" & vbCrLf
                sql &= " ,EXCAP1" & vbCrLf
                sql &= " ,EXCAP2" & vbCrLf
                sql &= " ,STYPE" & vbCrLf
                sql &= " ,NOHEALTHYN" & vbCrLf
                sql &= " ,APMARK" & vbCrLf
                sql &= " ,SENIORYN" & vbCrLf
                sql &= " ,OPITEMS2" & vbCrLf
                sql &= " ) Values " & vbCrLf
                sql &= " ('G'" & vbCrLf
                sql &= " ,'0002'" & vbCrLf
                sql &= " ,'57'" & vbCrLf
                sql &= " ,'" & APPLYYEAR & "'" & vbCrLf
                sql &= " ,''" & vbCrLf
                sql &= " ,'" & EXAMKIND & "'" & vbCrLf
                sql &= " ,'" & sGet_ExamLevels(EXAMLEVEL) & "'" & vbCrLf
                sql &= " ,'免填'" & vbCrLf
                sql &= " ,''" & vbCrLf
                sql &= " ,''" & vbCrLf
                sql &= " ,''" & vbCrLf
                sql &= " ,'" & IDNO & "'" & vbCrLf
                sql &= " ,'" & NAME & "'" & vbCrLf
                sql &= " ,'" & ENGNAME & "'" & vbCrLf
                sql &= " ,'" & SEX & "'" & vbCrLf

                sql &= " ,'" & BIRTHDAY & "'" & vbCrLf
                sql &= " ,''" & vbCrLf
                sql &= " ,'" & DEGREEID & "'" & vbCrLf
                sql &= " ,'" & GRADUATESTATUS & "'" & vbCrLf
                sql &= " ,'" & PHONED & "'" & vbCrLf
                sql &= " ,'" & PHONEN & "'" & vbCrLf
                sql &= " ,'" & CELLPHONE & "'" & vbCrLf
                sql &= " ,'" & EMAIL & "'" & vbCrLf
                sql &= " ,'" & ZIPCODE1 & "'" & vbCrLf
                sql &= " ,'" & ADDRESS1 & "'" & vbCrLf

                sql &= " ,'" & ZIPCODE2 & "'" & vbCrLf
                sql &= " ,'" & ADDRESS2 & "'" & vbCrLf
                sql &= " ,''" & vbCrLf
                sql &= " ,''" & vbCrLf
                sql &= " ,'G'" & vbCrLf
                sql &= " ,'N'" & vbCrLf
                sql &= " ,'N'" & vbCrLf
                sql &= " ,'N'" & vbCrLf
                sql &= " ,'非必填'" & vbCrLf
                sql &= " )" & vbCrLf

                Using OleCmd As New OleDb.OleDbCommand(sql, MyConn)
                    Try
                        If MyConn.State = ConnectionState.Closed Then MyConn.Open()
                        OleCmd.ExecuteNonQuery()
                        'If MyConn.State = ConnectionState.Open Then MyConn.Close()
                    Catch ex As Exception
                        If MyConn.State = ConnectionState.Open Then MyConn.Close()
                        Throw ex
                    End Try
                End Using
            Next
            If MyConn.State = ConnectionState.Open Then MyConn.Close()

        End Using

        'sql = "SELECT * FROM [Sheet1$]"
        'Dim objAdapter As OleDb.OleDbDataAdapter
        'If Not objAdapter Is Nothing Then
        '    objAdapter.Dispose()
        'End If
        'objAdapter = New OleDb.OleDbDataAdapter(sql, MyConn)
        'Dim objTable As New DataTable
        'objAdapter.Fill(objTable)
        'DataGrid_Temp.DataSource = objTable
        'DataGrid_Temp.DataBind()

        '將新建立的excel存入記憶體下載-----   Start
        'Dim strErrmsg As String = ""
        strErrmsg = ""
        Try
            Dim fr As New System.IO.FileStream(MyPath, IO.FileMode.Open)
            Dim br As New System.IO.BinaryReader(fr)
            Dim buf(fr.Length) As Byte

            fr.Read(buf, 0, fr.Length)
            fr.Close()

            Response.Clear()
            Response.ClearHeaders()
            Response.Buffer = True
            Response.AddHeader("content-disposition", "attachment;filename=" & HttpUtility.UrlEncode(MyFileName, System.Text.Encoding.UTF8))
            Response.ContentType = "Application/vnd.ms-Excel"
            'Common.RespWrite(Me, br.ReadBytes(fr.Length))
            Response.BinaryWrite(buf)
        Catch ex As Exception
            strErrmsg = ""
            strErrmsg += "無法存取該檔案!!!" & vbCrLf
            strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉) " & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
        Finally
            '刪除Temp中的資料
            If IO.File.Exists(MyPath) Then IO.File.Delete(MyPath)
            'If strErrmsg = "" Then Response.End()
            If strErrmsg = "" Then
                Call TIMS.CloseDbConn(objconn)
                Response.End()
            End If
            '將新建立的excel存入記憶體下載-----   End
        End Try
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
        End If

    End Sub

    '匯出
    Private Sub btnExport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport1.Click
        If tExamKind.Text = "" _
            OrElse hid_ExamTime.Value = "" _
            OrElse OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請先執行查詢作業!!")
            Exit Sub
        End If

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select ct.*" & vbCrLf
        '依 ocid,ExamKind,ExamTime
        sql &= " FROM CLASS_TECHEXAM ct" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and ct.ocid ='" & OCIDValue1.Value & "'"
        sql &= " and ct.ExamKind ='" & tExamKind.Text & "'"
        sql &= " and ct.ExamTime ='" & hid_ExamTime.Value & "'"
        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無有效資料!!")
            Exit Sub
        End If

        '匯出
        Call sUtl_Export1(OCIDValue1.Value, tExamKind.Text, hid_ExamTime.Value)
    End Sub


#Region "function 1"
#End Region

    '取得為第幾次檢定資料 
    Function sUtl_GetExamTime(ByVal ocid As String, ByVal examkind As String) As Integer
        Dim rst As Integer = 4  '可以輸入 1, 2, 3 (目前系統限定3種職類，以班為單位) 4:已超過系統3次

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select a.ExamKind" & vbCrLf
        sql &= " ,a.ExamTime" & vbCrLf
        '依 ocid 過濾,ExamKind 取得,ExamTime
        sql &= " FROM CLASS_TECHEXAM a" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and a.ocid ='" & ocid & "'" & vbCrLf
        sql &= " order by a.ExamTime" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        Dim ff3 As String = "ExamKind ='" & examkind & "'"
        Select Case dt.Rows.Count
            Case 0 '沒有資料
                rst = 1
            Case 1 '有1筆資料
                If dt.Select(ff3).Length > 0 Then
                    '找到(使用原來的次數)
                    rst = dt.Select(ff3)(0)("ExamTime")
                Else
                    '沒找到
                    rst = 2
                End If
            Case 2 '有2筆資料
                If dt.Select(ff3).Length > 0 Then
                    '找到(使用原來的次數)
                    rst = dt.Select(ff3)(0)("ExamTime")
                Else
                    '沒找到
                    rst = 3
                End If
            Case 3
                '有3筆資料
                If dt.Select(ff3).Length > 0 Then
                    '找到(使用原來的次數)
                    rst = dt.Select(ff3)(0)("ExamTime")
                    '沒找到'else rst=4(異常)
                End If
        End Select
        Return rst
    End Function

    '建立技檢對應值到 Clinet端
    Sub CreateJScript()
        Dim sql As String
        Dim dt As DataTable
        Dim JScript As String

        sql = "SELECT EXAMID,NAME FROM KEY_EXAM ORDER BY EXAMID"
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then Exit Sub

        Dim dr1 As DataRow = dt.Rows(0)
        Dim sEXAMID As String = Convert.ToString(dr1("EXAMID"))
        Dim sNAME As String = Convert.ToString(dr1("NAME"))
        vMsg = "範例：" & sEXAMID & "[" & sNAME & "]"
        TIMS.Tooltip(tExamKind, vMsg)

        JScript = "<script language='javascript'>" & vbCrLf
        JScript += "   function Get_Exam(ExamNo,obj){" & vbCrLf 'this Value (input), Name (output)
        JScript += "      var MyValue=ExamNo.value;" & vbCrLf
        JScript += "      var Exam=new Array;" & vbCrLf '設計陣列
        For Each dr As DataRow In dt.Rows
            JScript += "      Exam['" & dr("ExamID") & "']='" & dr("Name") & "';" & vbCrLf
        Next
        JScript += "      eval('document.form1.'+obj).value='';" & vbCrLf
        JScript += "      if(MyValue!=''){" & vbCrLf
        'JScript += "         eval('document.form1.'+obj).value='';" & vbCrLf
        JScript += "         if(Exam[MyValue]!=undefined){" & vbCrLf
        JScript += "            eval('document.form1.'+obj).value=Exam[MyValue];" & vbCrLf
        JScript += "         }" & vbCrLf
        JScript += "         else{" & vbCrLf
        JScript += "            alert('錯誤的代碼');" & vbCrLf
        JScript += "            ExamNo.focus();" & vbCrLf
        JScript += "         }" & vbCrLf
        JScript += "      }" & vbCrLf
        JScript += "   }" & vbCrLf
        JScript += "</script>" & vbCrLf
        JScript += "" & vbCrLf

        Page.RegisterStartupScript("Exam", JScript)
    End Sub

    '1.重新查詢
    Sub reloadSch1()
        tExamKind.Enabled = True

        btnSaveData1.CommandName = ""
        btnExport1.CommandName = ""
        btnBack1.CommandName = ""

        DataGrid1c.Visible = False
        DataGrid1b.Visible = False
        DataGrid1.Visible = False

        msg3.Text = ""
        msg2.Text = ""
        msg.Text = ""

        btnSaveData1.Visible = False '儲存
        btnExport1.Visible = False '匯出
        btnBack1.Visible = False '回上頁
    End Sub

    Function sUtl_GetStdTechExam(ByVal iOCID As Integer,
                                 ByVal sExamKind As String,
                                 ByVal iEXAMTIME As Integer) As DataTable
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT c.STXID" & vbCrLf
        sql &= " ,c.CTEID" & vbCrLf
        sql &= " ,c.SOCID" & vbCrLf
        sql &= " ,c.EXAMLEVEL" & vbCrLf
        sql &= " ,c.APPLYDATE" & vbCrLf
        sql &= " ,c.PASS" & vbCrLf
        sql &= " ,c.SENDOUTCERTDATE" & vbCrLf
        sql &= " ,c.EXAMNO" & vbCrLf
        sql &= " ,c.EXAMDATE" & vbCrLf
        sql &= " ,ct.EXAMTIME" & vbCrLf
        '依 ocid ,ExamKind ,ExamTime
        sql &= " FROM CLASS_TECHEXAM ct" & vbCrLf
        sql &= " JOIN STUD_TECHEXAM3 c on c.CTEID=ct.CTEID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " and ct.OCID =@OCID" & vbCrLf
        sql &= " and ct.ExamKind =@ExamKind" & vbCrLf
        sql &= " and ct.EXAMTIME =@EXAMTIME" & vbCrLf
        sql &= " ORDER BY c.EXAMLEVEL" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt1 As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.Int).Value = iOCID
            .Parameters.Add("ExamKind", SqlDbType.VarChar).Value = sExamKind
            .Parameters.Add("EXAMTIME", SqlDbType.Int).Value = iEXAMTIME
            dt1.Load(.ExecuteReader())
        End With
        Return dt1
    End Function

    Function Get_Examlevels1(ByRef dt1 As DataTable, ByVal socid As String, ByVal EXAMTIME As String) As String
        Dim rst As String = ""
        If dt1 Is Nothing Then Return rst
        If dt1.Rows.Count = 0 Then Return rst
        Dim ff3 As String = "SOCID=" & socid & " and EXAMTIME=" & EXAMTIME
        If dt1.Select(ff3).Length = 0 Then Return rst
        For Each dr1 As DataRow In dt1.Select(ff3)
            If rst <> "" Then rst &= ","
            rst &= Convert.ToString(dr1("EXAMLEVEL"))
        Next
        Return rst
    End Function

    Function Get_ApplyDate1(ByRef dt1 As DataTable, ByVal socid As String, ByVal EXAMTIME As String) As String
        Dim rst As String = ""
        If dt1 Is Nothing Then Return rst
        If dt1.Rows.Count = 0 Then Return rst
        Dim ff3 As String = "SOCID=" & socid & " and EXAMTIME=" & EXAMTIME
        If dt1.Select(ff3).Length = 0 Then Return rst
        rst = TIMS.Cdate3(dt1.Select(ff3)(0)("APPLYDATE"))
        Return rst
    End Function

#Region "NO USE"
    'Function Get_ExamPass1(ByRef dt1 As DataTable, ByVal socid As String, ByVal EXAMTIME As String) As String
    '    Dim rst As String = ""
    '    If dt1 Is Nothing Then Return rst
    '    If dt1.Rows.Count = 0 Then Return rst
    '    Dim ff3 As String = ""
    '    ff3 = "SOCID=" & socid & " and EXAMTIME=" & EXAMTIME & " and PASS='Y'"
    '    If dt1.Select(ff3).Length <> 0 Then Return "Y"
    '    ff3 = "SOCID=" & socid & " and EXAMTIME=" & EXAMTIME & " and PASS='N'"
    '    If dt1.Select(ff3).Length <> 0 Then Return "N"
    '    Return rst
    'End Function
#End Region

    Sub Get_ExamPass2(ByRef dt1 As DataTable, ByVal socid As String, ByVal EXAMTIME As String, ByRef cbo1 As CheckBoxList)
        If dt1 Is Nothing Then Exit Sub
        If dt1.Rows.Count = 0 Then Exit Sub
        Dim ff3 As String = ""
        'Dim strValue1 As String = ""
        With cbo1
            .Items.Clear()
            For ii As Integer = 1 To 5
                ff3 = "SOCID=" & socid & " and EXAMTIME=" & EXAMTIME & " and EXAMLEVEL=" & ii
                If dt1.Select(ff3).Length = 0 Then
                    Dim item1 As New ListItem(sGet_ExamLevel(ii), ii)
                    .Items.Add(item1)
                Else
                    'If strValue1 <> "" Then strValue1 &= ","
                    'strValue1 &= ii
                    Dim dr1 As DataRow = dt1.Select(ff3)(0)
                    Dim tmp1 As String = sGet_ExamLevel(ii) & "-" & TIMS.Cdate3(dr1("APPLYDATE"))
                    Dim item1 As New ListItem(tmp1, ii, False)
                    .Items.Add(item1)
                End If
            Next
        End With
        For Each listItem1 As ListItem In cbo1.Items
            If Not listItem1.Enabled Then listItem1.Selected = True
        Next
    End Sub

    '新增(單筆)
    Sub SaveData3i(ByVal eItem As DataGridItem, ByVal iCTEID As Integer, ByVal iSOCID As Integer,
                   ByVal sEXAMLEVELS2 As String, ByRef oConn As SqlConnection,
                   ByRef sCmd As SqlCommand, ByRef iCmd As SqlCommand,
                   ByRef uCmd As SqlCommand, ByRef dCmd As SqlCommand)
        'Dim myCheckbox2 As HtmlInputCheckBox = eItem.FindControl("Checkbox2") '全選 HtmlInputCheckBox
        'Dim hidSOCID As HtmlInputHidden = eItem.FindControl("hidSOCID") 'socid
        ''Dim hidSTEIDx As HtmlInputHidden = eItem.FindControl("hidSTEIDx")  'steid
        'Dim cbl1TL1 As CheckBoxList = eItem.FindControl("cbl1TrainLevel") '檢定級別
        'Dim cblVAL1 As String = "," & TIMS.GetChkBoxListValue(cbl1TL1) & ","
        Dim tApplyDate As TextBox = eItem.FindControl("tApplyDate")
        Call TIMS.OpenDbConn(oConn)
        For ii As Integer = 1 To 5
            Dim dt As New DataTable
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("CTEID", SqlDbType.Int).Value = iCTEID
                .Parameters.Add("SOCID", SqlDbType.Int).Value = iSOCID
                .Parameters.Add("EXAMLEVEL", SqlDbType.Int).Value = ii
                dt.Load(.ExecuteReader())
            End With
            ''停用刪除功能
            'If sEXAMLEVELS2.IndexOf("," & ii & ",") = -1 AndAlso dt.Rows.Count > 0 Then
            '    Dim dr1 As DataRow = dt.Rows(0)
            '    Dim iSTXID As Integer = dr1("STXID")
            '    With dCmd
            '        .Parameters.Clear()
            '        .Parameters.Add("STXID", SqlDbType.Int).Value = iSTXID
            '        .ExecuteNonQuery() 'dt.Load(.ExecuteReader())
            '    End With
            'End If
            If sEXAMLEVELS2.IndexOf("," & ii & ",") > -1 AndAlso dt.Rows.Count = 0 Then
                Dim iSTXID As Integer = DbAccess.GetNewId(objconn, "STUD_TECHEXAM3_STXID_SEQ,STUD_TECHEXAM3,STXID")
                With iCmd
                    .Parameters.Clear()
                    .Parameters.Add("STXID", SqlDbType.Int).Value = iSTXID
                    .Parameters.Add("CTEID", SqlDbType.Int).Value = iCTEID
                    .Parameters.Add("SOCID", SqlDbType.Int).Value = iSOCID
                    .Parameters.Add("EXAMLEVEL", SqlDbType.Int).Value = ii
                    .Parameters.Add("APPLYDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(tApplyDate.Text)
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    '.ExecuteNonQuery() 'dt.Load(.ExecuteReader())
                    DbAccess.ExecuteNonQuery(iCmd.CommandText, objconn, iCmd.Parameters)
                End With
            End If
        Next
    End Sub

    '修改(單筆)
    Sub SaveData3u(ByVal eItem As DataGridItem, ByVal iCTEID As Integer, ByVal iSOCID As Integer,
                   ByRef oConn As SqlConnection,
                   ByRef sCmd As SqlCommand, ByRef iCmd As SqlCommand,
                   ByRef uCmd As SqlCommand, ByRef dCmd As SqlCommand)
        Dim ddlExamPass As DropDownList = eItem.FindControl("ddlExamPass") '檢定結果
        Dim Textbox6 As TextBox = eItem.FindControl("Textbox6") '檢定日 Textbox6 (Img3)
        Dim TextBox4 As TextBox = eItem.FindControl("TextBox4") '製證日 TextBox4 (IMG2)
        Dim TextBox5 As TextBox = eItem.FindControl("TextBox5") '證號 TextBox5
        'Dim hidSOCID As HtmlInputHidden = eItem.FindControl("hidSOCID") 'socid
        Dim HidSTXID As HtmlInputHidden = eItem.FindControl("HidSTXID")
        'Dim myCheckbox2 As HtmlInputCheckBox = eItem.FindControl("Checkbox2") '全選 HtmlInputCheckBox
        'Dim hidSOCID As HtmlInputHidden = eItem.FindControl("hidSOCID") 'socid
        ''Dim hidSTEIDx As HtmlInputHidden = eItem.FindControl("hidSTEIDx")  'steid
        'Dim cbl1TL1 As CheckBoxList = eItem.FindControl("cbl1TrainLevel") '檢定級別
        'Dim cblVAL1 As String = "," & TIMS.GetChkBoxListValue(cbl1TL1) & ","
        'Dim tApplyDate As TextBox = eItem.FindControl("tApplyDate")
        'Call TIMS.OpenDbConn(oConn)
        Dim iSTXID As Integer = Val(HidSTXID.Value)
        Call TIMS.OpenDbConn(oConn)
        With uCmd
            .Parameters.Clear()
            .Parameters.Add("PASS", SqlDbType.Char).Value = IIf(TIMS.RstValue(ddlExamPass.SelectedValue, "Y,N") <> "", ddlExamPass.SelectedValue, Convert.DBNull)
            .Parameters.Add("SENDOUTCERTDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(TextBox4.Text)
            .Parameters.Add("EXAMNO", SqlDbType.VarChar).Value = IIf(Trim(TextBox5.Text) <> "", Trim(TextBox5.Text), Convert.DBNull)
            .Parameters.Add("EXAMDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(Textbox6.Text)
            .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID

            .Parameters.Add("STXID", SqlDbType.Int).Value = iSTXID
            .Parameters.Add("CTEID", SqlDbType.Int).Value = iCTEID
            .Parameters.Add("SOCID", SqlDbType.Int).Value = iSOCID
            '.ExecuteNonQuery() 'dt.Load(.ExecuteReader())
            DbAccess.ExecuteNonQuery(uCmd.CommandText, objconn, uCmd.Parameters)
        End With
    End Sub

    '單引號問題修正
    Function ReplaceText(ByVal MyText As String) As String
        'MyText = Replace(MyText, "'", "''")
        If MyText Is Nothing Then MyText = ""
        MyText = TIMS.ClearSQM(MyText)
        Return MyText
    End Function

    '存檔(技能檢定)
    Sub SaveData2()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " UPDATE STUD_TECHEXAM3" & vbCrLf
        sql &= " SET PASS=@PASS" & vbCrLf
        sql &= " ,SENDOUTCERTDATE=@SENDOUTCERTDATE" & vbCrLf
        sql &= " ,EXAMNO=@EXAMNO" & vbCrLf
        sql &= " ,EXAMDATE=@EXAMDATE" & vbCrLf
        sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE=getdate()" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and STXID=@STXID" & vbCrLf
        sql &= " and CTEID=@CTEID" & vbCrLf
        sql &= " and SOCID=@SOCID" & vbCrLf
        Dim uCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql &= " insert into STUD_TECHEXAM3(STXID,CTEID,SOCID,EXAMLEVEL,APPLYDATE,MODIFYACCT,MODIFYDATE)" & vbCrLf
        sql &= " values(@STXID,@CTEID,@SOCID,@EXAMLEVEL,@APPLYDATE,@MODIFYACCT,getdate())" & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql &= " SELECT STXID FROM STUD_TECHEXAM3" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and CTEID=@CTEID AND SOCID=@SOCID AND EXAMLEVEL=@EXAMLEVEL" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql &= " DELETE STUD_TECHEXAM3 WHERE STXID=@STXID" & vbCrLf
        Dim dCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql &= " insert into CLASS_TECHEXAM(CTEID,OCID,EXAMKIND,EXAMTIME,MODIFYACCT,MODIFYDATE)" & vbCrLf
        sql &= " values(@CTEID,@OCID,@EXAMKIND,@EXAMTIME,@MODIFYACCT,getdate())" & vbCrLf
        Dim iCmd2 As New SqlCommand(sql, objconn)

        'Select Case btnSaveData1.CommandName
        '    Case cst_申請設定
        '    Case cst_結果輸入
        'End Select
        Call TIMS.OpenDbConn(objconn)
        Select Case btnSaveData1.CommandName
            Case cst_申請設定
                Dim iCTEID As Integer = 0
                If hid_CTEID.Value = "" Then
                    iCTEID = DbAccess.GetNewId(objconn, "CLASS_TECHEXAM_CTEID_SEQ,CLASS_TECHEXAM,CTEID")
                    With iCmd2
                        .Parameters.Clear()
                        .Parameters.Add("CTEID", SqlDbType.Int).Value = iCTEID
                        .Parameters.Add("OCID", SqlDbType.Int).Value = Val(hid_OCID.Value)
                        .Parameters.Add("EXAMKIND", SqlDbType.VarChar).Value = hid_EXAMKIND.Value
                        .Parameters.Add("EXAMTIME", SqlDbType.VarChar).Value = Val(hid_ExamTime.Value)
                        .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                        '.ExecuteNonQuery()
                        DbAccess.ExecuteNonQuery(iCmd2.CommandText, objconn, iCmd2.Parameters)
                    End With
                    hid_CTEID.Value = iCTEID
                End If
            Case cst_結果輸入

        End Select

        Select Case btnSaveData1.CommandName
            Case cst_申請設定
                Dim iSave As Integer = 0
                For Each eItem As DataGridItem In DataGrid1b.Items
                    Dim myCheckbox2 As HtmlInputCheckBox = eItem.FindControl("Checkbox2") '全選 HtmlInputCheckBox
                    Dim hidSOCID As HtmlInputHidden = eItem.FindControl("hidSOCID") 'socid
                    'Dim hidSTEIDx As HtmlInputHidden = eItem.FindControl("hidSTEIDx")  'steid
                    Dim cbl1TL1 As CheckBoxList = eItem.FindControl("cbl1TrainLevel") '檢定級別
                    Dim sCblVAL1 As String = TIMS.GetChkBoxListValue2(cbl1TL1)
                    If sCblVAL1 <> "" Then sCblVAL1 = "," & sCblVAL1 & "," '(若不為空值左右加逗號)

                    Dim tApplyDate As TextBox = eItem.FindControl("tApplyDate")
                    If myCheckbox2.Checked AndAlso sCblVAL1 <> "" _
                        AndAlso hid_CTEID.Value <> "" AndAlso hidSOCID.Value <> "" Then
                        iSave += 1
                        '新增(修改)
                        Call SaveData3i(eItem, hid_CTEID.Value, hidSOCID.Value, sCblVAL1,
                                        objconn, sCmd, iCmd, uCmd, dCmd)
                    End If
                Next

                If iSave = 0 Then
                    Common.MessageBox(Me, "無有效資料存取!!")
                    Exit Sub
                End If

                Common.MessageBox(Me, "儲存成功")
                ' Common.MessageBox(Me, "查詢")
                Call reloadSch1()
                '查詢
                Call show_dg1(OCIDValue1.Value, tExamKind.Text)

            Case cst_結果輸入
                '檢查輸入資料長度。
                For Each eItem As DataGridItem In DataGrid1c.Items
                    '檢定結果
                    Dim ddlExamPass As DropDownList = eItem.FindControl("ddlExamPass")
                    '檢定日 Textbox6 (Img3)
                    '製證日 TextBox4 (IMG2)
                    '證號 TextBox5
                    Dim Textbox6 As TextBox = eItem.FindControl("Textbox6")
                    Dim TextBox4 As TextBox = eItem.FindControl("TextBox4")
                    Dim TextBox5 As TextBox = eItem.FindControl("TextBox5")
                    Dim hidSOCID As HtmlInputHidden = eItem.FindControl("hidSOCID") 'socid
                    Dim HidSTXID As HtmlInputHidden = eItem.FindControl("HidSTXID")

                    'Dim dt As DataTable = New DataTable
                    If HidSTXID.Value <> "" Then
                        '有序號使用UPDATE
                        '新增(UPDATE)
                        Select Case ddlExamPass.SelectedValue
                            Case "Y"
                            Case "N"
                            Case Else
                        End Select
                        Textbox6.Text = TIMS.ClearSQM(Textbox6.Text)
                        TextBox4.Text = TIMS.ClearSQM(TextBox4.Text)
                        If Textbox6.Text <> "" AndAlso Not TIMS.IsDate1(Textbox6.Text) Then
                            Common.MessageBox(Me, "檢定日 日期格式有誤，請檢查輸入資料!(yyyy/MM/dd)")
                            Exit Sub
                        End If
                        If TextBox4.Text <> "" AndAlso Not TIMS.IsDate1(TextBox4.Text) Then
                            Common.MessageBox(Me, "製證日 日期格式有誤，請檢查輸入資料!(yyyy/MM/dd)")
                            Exit Sub
                        End If
                        TextBox5.Text = TIMS.ClearSQM(TextBox5.Text)
                        If TextBox5.Text <> "" Then
                            If Len(TextBox5.Text) > 20 Then
                                Common.MessageBox(Me, "證號長度超過系統範圍20，請檢查輸入資料!!")
                                Exit Sub
                            End If
                        End If
                    End If
                Next

                For Each eItem As DataGridItem In DataGrid1c.Items
                    'Dim ddlExamPass As DropDownList = eItem.FindControl("ddlExamPass")  '檢定結果
                    'Dim Textbox6 As TextBox = eItem.FindControl("Textbox6") '檢定日 Textbox6 (Img3)
                    'Dim TextBox4 As TextBox = eItem.FindControl("TextBox4") '製證日 TextBox4 (IMG2)
                    'Dim TextBox5 As TextBox = eItem.FindControl("TextBox5") '證號 TextBox5
                    Dim hidSOCID As HtmlInputHidden = eItem.FindControl("hidSOCID") 'socid
                    Dim HidSTXID As HtmlInputHidden = eItem.FindControl("HidSTXID")

                    Dim dt As DataTable = New DataTable
                    If HidSTXID.Value <> "" AndAlso hid_CTEID.Value <> "" AndAlso hidSOCID.Value <> "" Then
                        '有序號使用UPDATE '新增(修改)
                        Call SaveData3u(eItem, hid_CTEID.Value, hidSOCID.Value,
                                        objconn, sCmd, iCmd, uCmd, dCmd)

                    End If
                Next
                'Call TIMS.CloseDbConn(da1.UpdateCommand.Connection)
                Common.MessageBox(Me, "儲存成功")

                ' Common.MessageBox(Me, "查詢")
                Call reloadSch1()
                '查詢
                Call show_dg1(OCIDValue1.Value, tExamKind.Text)
        End Select
    End Sub

    Function sGet_ExamLevels(ByVal str1 As String) As String
        Dim rst As String = ""
        Dim astr1 As String() = str1.Split(",")
        For Each str2 As String In astr1
            If rst <> "" Then rst &= ","
            rst &= sGet_ExamLevel(str2)
        Next
        Return rst
    End Function

    Function sGet_ExamLevel(ByVal str1 As String) As String
        Dim rst As String = ""
        Dim slevel As String = ""
        Dim iINX As Integer = cst_xExamLevelv1.IndexOf(str1)
        If iINX = -1 Then Return rst
        Dim aEl As String() = cst_xExamLevelv1.Split(",")
        For ii As Integer = 0 To aEl.Length - 1
            If aEl(ii) = str1 Then
                slevel = cst_xExamLeveln1.Split(",")(ii)
                Return slevel
            End If
        Next
        Return rst


        'Select Case str1
        '    Case "1"
        '        slevel = cst_xExamLeveln1.Split(",")(0)
        '    Case "2"
        '        slevel = "乙級"
        '    Case "3"
        '        slevel = "丙級"
        '    Case "4"
        '        slevel = "單一級"
        '    Case "5"
        '        slevel = "不分級"
        'End Select
        'Return slevel
    End Function

    Private Sub btnGETvalue1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGETvalue1.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub

    Private Sub btnSetOneOCID_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSetOneOCID.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    '保留搜尋值
    Sub keepSearch1()
        Dim sch1 As String = ""
        TIMS.SetMyValue(sch1, "tExamKind", tExamKind.Text)
        TIMS.SetMyValue(sch1, "tExamName", tExamName.Text)
        TIMS.SetMyValue(sch1, "hid_ExamTime", hid_ExamTime.Value)
        Session(cst_ss_Sd07001_sch1) = sch1
    End Sub

    '回上頁
    Private Sub btnBack1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack1.Click
        tExamKind.Text = ""
        tExamName.Text = ""
        hid_ExamTime.Value = ""
        If Convert.ToString(Session(cst_ss_Sd07001_sch1)) Is Nothing Then
            Dim sch1 As String = Session(cst_ss_Sd07001_sch1)
            tExamKind.Text = TIMS.GetMyValue(sch1, "tExamKind")
            tExamName.Text = TIMS.GetMyValue(sch1, "tExamName")
            hid_ExamTime.Value = TIMS.GetMyValue(sch1, "hid_ExamTime")
            Session(cst_ss_Sd07001_sch1) = Nothing
        End If

        Call reloadSch1()
        '查詢
        Call show_dg1(OCIDValue1.Value, tExamKind.Text)
    End Sub

    '存檔(技能檢定)
    Private Sub btnSaveData1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveData1.Click
        '存檔(技能檢定)
        Call SaveData2()
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
