Partial Class SD_05_034_R
    Inherits AuthBasePage

    'SD_05_002_R_Rpt.aspx                                     '出缺勤明細表
    'excuse_list                                              '請假、缺曠課累計時數統計表
    Dim strTPlanIDs As String = ""
    'Const cst_printR1 As String = "SD_05_034_R1"             '出缺勤明細表
    'SD_05_034_R_Rpt.aspx
    Const cst_printR1aspx As String = "SD_05_034_R_Rpt.aspx"  '屆退官兵出缺勤明細表
    'Const cst_printFN2 As String = "SD_05_034_R2"            '(請假缺曠課累計時數統計表)國軍屆退官兵參訓學員請假、缺曠課統計表
    Const cst_printFN3 As String = "SD_05_034_R3"             '(請假缺曠課累計時數統計表)國軍屆退官兵參訓學員請假、缺曠課統計表

    'Dim flagYear2017 As Boolean = False
    'flagYear2017 = TIMS.Get_UseLEAVE_2017(Me)

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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        'Dim flagYear2017 As Boolean = False
        'flagYear2017 = TIMS.Get_UseLEAVE_2017(Me)

        If Not IsPostBack Then Call create1()

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
    End Sub

    '第1次網頁載入
    Sub create1()
        '複選，選項有：自辦職前訓練、產訓合作訓練(自辦)、主題產業職業訓練(職前)、客制化訓用合一。
        '02:自辦職前訓練
        '14:產訓合作訓練(自辦)
        '64:主題產業職業訓練(職前)
        '65:客製化訓用合一 
        'Const Cst_TPlanID02Plan2 As String = "'02','14','64','65'"
        'Public Const Cst_TPlanID02Plan2 As String = "02,14,64,65"
        cblTPlanID = TIMS.Get_TPlan(cblTPlanID, , 1, "Y", "TPlanID IN (" & TIMS.Cst_TPlanID02Plan2 & ")", objconn)
        cblTPlanID.Items(0).Selected = True

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT CONVERT(VARCHAR, dbo.TRUNC_DATETIME(GETDATE()), 111) + '/01' STDATE1 " & vbCrLf
        'sql &= "        ,CONVERT(VARCHAR, LAST_DAY(GETDATE()), 111) STDATE2 " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, EOMONTH(GETDATE()), 111) STDATE2 " & vbCrLf  'edit，by:20181101
        sql &= " " & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable

        With sCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            Dim dr1 As DataRow = dt.Rows(0)
            'Dim sTODAY As String = TIMS.GetSysDate(objconn)
            STDate1.Text = TIMS.Cdate3(dr1("STDATE1"))
            STDate2.Text = TIMS.Cdate3(dr1("STDATE2"))
        End If

        btnPrint1.Attributes("onclick") = "javascript:return print();"

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        hbtnOrg.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        HistoryRID.Visible = True
        center.Enabled = True
        hbtnOrg.Visible = True
        If sm.UserInfo.DistID <> "000" Then
            '非署(局)鎖定分署(轄區)。
            HistoryRID.Visible = False
            hbtnOrg.Visible = False
            center.Enabled = False '.Text 
        End If
    End Sub

    '檢核
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        '選擇計畫
        Dim objCbl As CheckBoxList = Me.cblTPlanID
        Dim flagAll As Boolean = False '是否選擇全部
        For Each objLItem As ListItem In objCbl.Items
            If objLItem.Selected AndAlso objLItem.Value = "" Then
                '全部
                flagAll = True '(選擇全部)
                Exit For '離開此迴圈
            End If
            If objLItem.Value <> "" Then Exit For '有值離開此迴圈
        Next
        Dim itemstr As String = ""
        For Each objLItem As ListItem In objCbl.Items
            If objLItem.Value <> "" Then
                If flagAll Then
                    '只要有值就加入
                    If itemstr <> "" Then itemstr += ","
                    'itemstr += "'" & objLItem.Value & "'"
                    itemstr += objLItem.Value
                Else
                    '要有勾選才能加
                    If objLItem.Selected Then
                        If itemstr <> "" Then itemstr += ","
                        'itemstr += "'" & objLItem.Value & "'"
                        itemstr += objLItem.Value
                    End If
                End If
            End If
        Next
        If itemstr = "" Then Errmsg += "訓練計畫 必須選擇" & vbCrLf
        If Errmsg = "" Then strTPlanIDs = itemstr

        STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        STDate1.Text = TIMS.Cdate3(STDate1.Text)
        STDate2.Text = TIMS.ClearSQM(STDate2.Text)
        STDate2.Text = TIMS.Cdate3(STDate2.Text)

#Region "(No Use)"

        'If Trim(start_date.Text) <> "" Then start_date.Text = Trim(start_date.Text) Else start_date.Text = ""
        'If Trim(end_date.Text) <> "" Then end_date.Text = Trim(end_date.Text) Else end_date.Text = ""

        'If start_date.Text <> "" Then
        '    If Not TIMS.IsDate1(start_date.Text) Then
        '        Errmsg += "時間區間 起始日期格式有誤" & vbCrLf
        '    End If
        '    If Errmsg = "" Then
        '        start_date.Text = CDate(start_date.Text).ToString("yyyy/MM/dd")
        '    End If
        'Else
        '    'Errmsg += "時間區間 起始日期 為必填" & vbCrLf
        'End If

        'If end_date.Text <> "" Then
        '    If Not TIMS.IsDate1(end_date.Text) Then
        '        Errmsg += "時間區間 迄止日期格式有誤" & vbCrLf
        '    End If
        '    If Errmsg = "" Then
        '        end_date.Text = CDate(end_date.Text).ToString("yyyy/MM/dd")
        '    End If
        'Else
        '    'Errmsg += "時間區間 迄止日期 為必填" & vbCrLf
        'End If

#End Region

        If STDate1.Text = "" OrElse STDate2.Text = "" Then Errmsg += "時間區間 為必填" & vbCrLf

        If Errmsg = "" Then
            If STDate1.Text.ToString <> "" AndAlso STDate2.Text.ToString <> "" Then
                If DateDiff(DateInterval.Day, CDate(STDate1.Text), CDate(STDate2.Text)) < 0 Then
                    Dim sTmp As String = STDate1.Text
                    STDate1.Text = STDate2.Text
                    STDate2.Text = sTmp
                    'Errmsg += "【時間區間】的起日 不得大於 迄日!!" & vbCrLf
                End If
            End If
        End If

        Select Case RblPrintType1.SelectedValue
            Case "1", "2"
            Case Else
                '列印格式
                Errmsg += "請選擇【列印格式】!!" & vbCrLf
        End Select

        'If Convert.ToString(OCID.SelectedValue) = "" OrElse Not IsNumeric(OCID.SelectedValue) Then Errmsg += "統計對象 為必選" & vbCrLf

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '列印
    Protected Sub btnPrint1_Click(sender As Object, e As EventArgs) Handles btnPrint1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            '檢核有誤。
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        'Dim querystr As String = ""
        Dim RID As String = sm.UserInfo.RID
        If Me.RIDValue.Value <> "" Then RID = Me.RIDValue.Value

        'If Hiditem2.Value = "0" OrElse Hiditem2.Value = "" Then Hiditem2.Value = "1"
        'If Hiditem4.Value = "0" OrElse Hiditem4.Value = "" Then Hiditem4.Value = "1"

        Dim querystr As String = ""
        querystr &= "&TPlanID=" & strTPlanIDs 'sm.UserInfo.TPlanID
        querystr &= "&RID=" & RID
        querystr &= "&STDate1=" & STDate1.Text
        querystr &= "&STDate2=" & STDate2.Text
        querystr &= "&UserID=" & sm.UserInfo.UserID
        '<![CDATA[ ]]>
        Select Case RblPrintType1.SelectedValue
            Case "1"
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printR1, querystr)
                '屆退官兵出缺勤明細表
                ReportQuery.Redirect(Me, cst_printR1aspx, querystr)
            Case "2"
                '請假、缺曠課累計時數統計表
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, querystr)
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN3, querystr)
                'Dim flagYear2017 As Boolean = False
                'flagYear2017 = TIMS.Get_UseLEAVE_2017(Me)
                'Dim sPrintName1 As String = cst_printFN2
                'If flagYear2017 Then sPrintName1 = cst_printFN3
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sPrintName1, querystr)
        End Select
    End Sub
End Class