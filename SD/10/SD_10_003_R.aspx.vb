Partial Class SD_10_003_R
    Inherits AuthBasePage

    Const cst_printFN1s As String = "close" '2012年前舊'自辦 old
    Const cst_printFN1o As String = "close_1" '2012年前舊'委訓 old
    Const cst_printFN2c As String = "close_2c" '大於40筆資料用雙列顯示 (背面)
    Const cst_printFN2b As String = "close_2b" '小於n筆資料用單列顯示 (背面)

    Const cst_printFN2s As String = "close_21" '2012年後新'自辦 (正面)
    Const cst_printFN2o As String = "close_22" '2012年後新'委訓 (正面)
    'Const cst_printFN6s As String = "close21_16" '2016年後新'自辦 (正面)
    'Const cst_printFN6o As String = "close22_16" '2016年後新'委訓 (正面)

    '正面'2012年前舊
    'close'自辦 'close_1'委訓
    '正面'2012年後新
    'close_21 '自辦 'close_22 '委訓
    '2016 針對特定計畫:職前訓練(02','14','17','20','21','26','34','37','47','53','55','58','59','61','62','64','65')共計17支計畫，修改在訓證明、受訓證明、結訓證明等3張表件
    'close21_16 'close22_16

    '反面
    'close_2c '大於40筆資料用雙列顯示 'close_2b '小於n筆資料用單列顯示

    '0 ' 40 '小於n(40)筆資料用單列顯示
    Const Cst_Max_classCnt As Integer = 30 '40

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        If Not IsPostBack Then
            CCreate1()
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", , "search")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

    End Sub

    Sub CCreate1()
        msg.Text = ""
        search.Attributes("onclick") = "javascript:return search1();"
        submit.Attributes("onclick") = "javascript:return chkCertificateNo();"
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        hidSearchTag.Value = ""
        'trPrintStyle2.Visible = False
        If sm.UserInfo.Years >= 2012 Then Common.SetListItem(PrintStyle3, "2") '使用新版
        '結訓'證明字號
        ProveNum.Text = TIMS.GetGlobalVar(Me, "11", "1", objconn)
    End Sub
    Private Sub DG_stud_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_stud.ItemDataBound
        Const Cst_學員狀態 As Integer = 4
        'Dim Checkbox1 As HtmlInputCheckBox

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim oCheckbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                oCheckbox1.Value = Convert.ToString(drv("StudentID")) '.ToString
                Dim STUDSTATUS_N As String = TIMS.GET_STUDSTATUS_N(drv("StudStatus"))
                e.Item.Cells(Cst_學員狀態).Text = STUDSTATUS_N '"在訓"
                'CreditPoints-是否核發結訓證書 1:是／0:否
                Dim vCreditPoints As String = Convert.ToString(drv("CreditPoints"))

                Select Case $"{drv("StudStatus")}"
                    Case "2"
                        oCheckbox1.Disabled = True
                        TIMS.Tooltip(oCheckbox1, "(離訓)", True)
                    Case "3"
                        oCheckbox1.Disabled = True
                        TIMS.Tooltip(oCheckbox1, "(退訓)", True)
                End Select
                '自辦、區域、企委,06,07,70,
                'If ",06,07,70,".IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '    If Not oCheckbox1.Disabled AndAlso vCreditPoints <> "1" Then
                '        oCheckbox1.Disabled = True
                '        TIMS.Tooltip(oCheckbox1, "請確認【是否核發結訓證書】為是", True)
                '    End If
                'End If
        End Select
    End Sub

    '單一班級
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        DG_stud.Visible = False
        submit.Visible = False
    End Sub

    '查詢
    Private Sub search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search.Click
        Dim iClassCnt As Integer = 0

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        'CREATE INDEX IX_MVIEW_CLASS_SCHEDULE_OCID ON MVIEW_CLASS_SCHEDULE (OCID);
        '該班 訓練課程與授課時數資料(實際排課授課時數)
        If OCIDValue1.Value <> "" Then
            'MVIEW_CLASS_SCHEDULE (VIEW_CLASS_SCHEDULE) /CLASS_SCHEDULE/COURSE_COURSEINFO/TEACH_TEACHERINFO
            Dim pms_1 As New Hashtable From {{"OCID", TIMS.CINT1(OCIDValue1.Value)}}
            Dim sql As String = ""
            sql &= " SELECT DISTINCT SC.OCID,SC.COURSEID,SC.COURSENAME" & vbCrLf
            sql &= " FROM MVIEW_CLASS_SCHEDULE sc " & vbCrLf
            sql &= " WHERE sc.OCID=@OCID"
            Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, pms_1)
            If TIMS.dtHaveDATA(dt) Then iClassCnt = dt.Rows.Count
        End If
        If iClassCnt = 0 Then
            'MVIEW_CLASS_SCHEDULE (VIEW_CLASS_SCHEDULE)
            Dim strMsg As String = String.Concat("該班無訓練課程與授課時數資料!!", vbCrLf, "(若是剛剛才新增，因系統效能問題，請稍後再試)", vbCrLf)
            Common.MessageBox(Me, strMsg)
            Exit Sub
        End If

        Dim pms1_1 As New Hashtable From {{"OCID", TIMS.CINT1(OCIDValue1.Value)}, {"TPLANID", sm.UserInfo.TPlanID}}

        Dim sqlstr1 As String = ""
        sqlstr1 &= " select b.studentid,c.name,c.EngName,b.OCID,b.StudStatus,b.CreditPoints " & vbCrLf
        sqlstr1 &= " ,'" & iClassCnt & "' classCnt" & vbCrLf
        sqlstr1 &= " from class_classinfo a" & vbCrLf
        sqlstr1 &= " join class_studentsofclass b on a.ocid=b.ocid" & vbCrLf
        sqlstr1 &= " join stud_studentinfo c on b.sid=c.sid" & vbCrLf
        sqlstr1 &= " join id_plan ip on ip.planid=a.planid" & vbCrLf
        sqlstr1 &= " where b.OCID=@OCID" & vbCrLf
        sqlstr1 &= " and ip.TPLANID=@TPLANID" & vbCrLf
        If sm.UserInfo.LID > 0 Then
            pms1_1.Add("RID", sm.UserInfo.RID)
            sqlstr1 &= " and a.RID=@RID" & vbCrLf
        End If
        sqlstr1 &= " order by b.StudentID" & vbCrLf
        Dim stud_table As DataTable = DbAccess.GetDataTable(sqlstr1, objconn, pms1_1)

        DG_stud.Visible = False
        msg.Text = "查無資料!!"
        submit.Visible = False

        If TIMS.dtNODATA(stud_table) Then Return

        DG_stud.Visible = True
        msg.Text = ""
        submit.Visible = True
        'DG_stud.Visible = True
        'msg.Visible = False
        DG_stud.DataSource = stud_table
        DG_stud.DataBind()
    End Sub

    '送出
    Private Sub Submit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles submit.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        ProveNum.Text = TIMS.ClearSQM(ProveNum.Text) '結訓證書字號

        Dim StudentID As String = ""
        ViewState("classCnt") = ""
        For Each eItem As DataGridItem In DG_stud.Items
            Dim Checkbox1 As HtmlInputCheckBox = CType(eItem.FindControl("Checkbox1"), HtmlInputCheckBox)
            Dim classCnt As HtmlInputHidden = CType(eItem.FindControl("classCnt"), HtmlInputHidden)
            If ViewState("classCnt") = "" AndAlso classCnt.Value <> "" Then ViewState("classCnt") = classCnt.Value
            If Checkbox1.Checked AndAlso Checkbox1.Value <> "" Then
                '當被選取要做的事情 
                Checkbox1.Value = TIMS.ClearSQM(Checkbox1.Value)
                StudentID &= String.Concat(If(StudentID <> "", ",", ""), "\'", Checkbox1.Value, "\'")
            End If
        Next

        If StudentID = "" Then
            Common.MessageBox(Me, "請先勾選學員。")
            Exit Sub
        End If

        Select Case PrintStyle2.SelectedValue
            Case "1" '正面
            Case Else '反面
                Select Case PrintStyle3.SelectedValue ' 列印版本 (1:2012前 2:2012後)
                    Case "1" '2012年前舊
                        Common.MessageBox(Me, "舊版本無反面設計，請重新選擇。")
                        Exit Sub
                End Select
        End Select

        Dim MyValue As String = ""
        MyValue &= "&DistID=" & sm.UserInfo.DistID
        MyValue &= "&StudentID=" & StudentID
        MyValue &= "&OCID=" & OCIDValue1.Value
        MyValue &= "&ProveNum=" & Convert.ToString(ProveNum.Text)
        MyValue &= "&Type=2" & OCIDValue1.Value
        'MyValue &= "&YearType=" & rblYearType1.SelectedValue
        MyValue &= "&rblYearType1=" & rblYearType1.SelectedValue

        'Dim w16R1 As String = TIMS.Utl_GetConfigSet("w16R1") 'Y
        Dim RptFN1 As String = ""
        Select Case PrintStyle3.SelectedValue ' 列印版本
            Case "1" '2012年前舊
                Select Case PrintStyle.SelectedValue
                    Case "1"
                        RptFN1 = cst_printFN1s '自辦
                    Case Else
                        RptFN1 = cst_printFN1o '委訓
                End Select

            Case Else '2012年後新
                Select Case PrintStyle2.SelectedValue ' 列印版面
                    Case "1" '正面
                        Select Case PrintStyle.SelectedValue
                            Case "1"
                                RptFN1 = cst_printFN2s '自辦
                                'If w16R1 = "Y" Then RptFN1 = cst_printFN6s
                            Case Else
                                RptFN1 = cst_printFN2o '委訓
                                'If w16R1 = "Y" Then RptFN1 = cst_printFN6o
                        End Select

                    Case Else '反面
                        Dim flagType As Integer = 1 '1:單行列印 2:雙行列印
                        flagType = 1
                        If Cst_Max_classCnt <> 0 Then
                            '小於n筆資料用單列顯示
                            If CInt(Me.ViewState("classCnt")) > Cst_Max_classCnt Then flagType = 2
                        End If

                        Select Case flagType
                            Case 2
                                RptFN1 = cst_printFN2c '大於40筆資料用雙列顯示
                            Case Else '1
                                RptFN1 = cst_printFN2b '小於n筆資料用單列顯示
                        End Select
                End Select
        End Select

        If RptFN1 = "" Then
            Common.MessageBox(Me, "報表設定不正確，請重新選擇。")
            Exit Sub
        End If
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, RptFN1, MyValue)
    End Sub
End Class
