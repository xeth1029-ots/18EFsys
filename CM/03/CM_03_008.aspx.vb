Partial Class CM_03_008
    Inherits AuthBasePage

#Region "HIDE"
    'Key_RejectTReason
    'ReportQuery CM_03_008*.jrxml
    'CM_03_008_b  '統計資料'報表
    'CM_03_008  '明細資料'報表
    '2016
    'CM_03_008_b_2  '統計資料'報表
    'CM_03_008_2  '明細資料'報表
    'OLD
    'Const cst_printFN1A As String = "CM_03_008" '明細資料 "CM_03_008"
    'Const cst_printFN1B As String = "CM_03_008_b" '統計資料 "CM_03_008_b"
    '2016 CM_03_008*.jrxml
    'Const cst_printFN2B As String = "CM_03_008_2_B" '統計資料 "CM_03_008_b_2"/CM_03_008_2_B
    'Const cst_printFN2A As String = "CM_03_008_2" '明細資料 "CM_03_008_2" CM_03_008_2*.jrxml

    'Dim str_report_2 As String = "" '統計資料 "CM_03_008_b"
    'Dim str_report_1 As String = "" '明細資料 "CM_03_008"

    '奉召服兵役->自願、接受徵集入營者 CM_03_008_3*.jrxml
    'Const cst_printFN2C As String = "CM_03_008_3_C" '統計資料 (年度為空)
    'Const cst_printFN2B As String = "CM_03_008_3_B" '統計資料 "CM_03_008_b_2"/CM_03_008_2_B
    'Const cst_printFN2A As String = "CM_03_008_3" '明細資料 "CM_03_008_2" CM_03_008_2*.jrxml
#End Region

    '奉召服兵役->自願、接受徵集入營者 CM_03_008_4*.jrxml
    Const cst_printFN2C As String = "CM_03_008_4_C" '統計資料 (年度為空)
    Const cst_printFN2B As String = "CM_03_008_4_B" '統計資料 
    Const cst_printFN2A As String = "CM_03_008_4" '明細資料 

    Const cst_else_1 As String = "_else_1" '統計資料 (2015之前)
    Const cst_else_2 As String = "_else_2" '明細資料 (2015之前)

    Const cst_2016_1 As String = "_2016_1" '統計資料 (2016)
    Const cst_2016_2 As String = "_2016_2" '明細資料 (2016)

    Const cst_2017_1 As String = "_2017_1" '統計資料 (2017)
    Const cst_2017_2 As String = "_2017_2" '明細資料 (2017)

    Const cst_2020_1 As String = "_2020_1" '統計資料 (2020)
    Const cst_2020_2 As String = "_2020_2" '明細資料 (2020)

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        'str_report_2 = cst_printFN1B '統計資料 "CM_03_008_b"
        'str_report_1 = cst_printFN1A '明細資料 "CM_03_008"
        'If sm.UserInfo.Years >= "2016" Then '2016
        '    str_report_2 = cst_printFN2B '統計資料 "CM_03_008_b"
        '    str_report_1 = cst_printFN2A '明細資料 "CM_03_008"
        'End If
        'str_report_2 = cst_printFN2B '統計資料 "CM_03_008_2_B"
        'str_report_1 = cst_printFN2A '明細資料 "CM_03_008_2"

        If Not IsPostBack Then
            msg.Text = ""
            ExportMsg.Text = ""
            DataGrid1.Visible = False

            Call CreateItem()

            OCID.Style("display") = "none"
            msg.Text = TIMS.cst_NODATAMsg11

            DistID.Attributes("onclick") = "ClearData();"
            TPlanID.Attributes("onclick") = "ClearData();"

            '選擇全部轄區
            DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
            '選擇全部訓練計畫
            TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
            '選擇全部身分別
            Identity.Attributes("onclick") = "SelectAll('Identity','Identity_List');"
            '列印檢查
            Print.Attributes("onclick") = "javascript:return CheckPrint();"

            Button3.Style("display") = "none"
        End If

    End Sub

    Sub CreateItem()
        FTDate1.Text = TIMS.Cdate3(Now.Year.ToString() & "/1/1")
        FTDate2.Text = TIMS.Cdate3(Now.Date)

        '年度
        Syear = TIMS.GetSyear(Syear)
        Common.SetListItem(Syear, sm.UserInfo.Years)

        '轄區
        DistID = TIMS.Get_DistID(DistID)
        DistID.Items.Insert(0, New ListItem("全部", ""))
        '計畫
        TPlanID = TIMS.Get_TPlan(TPlanID, TIMS.dtNothing(), 1, "Y")
        '學員身份
        Identity = TIMS.Get_Identity(Identity, 31, objconn) ' 2011/02/24 改成全部

        '預算別 BUDID IN ('01','02','03')
        BudID = TIMS.Get_Budget(BudID, 38, objconn)

        'Dim dt As DataTable
        'Dim sql As String = ""
        'sql = "" & vbCrLf
        'sql &= " SELECT BUDID,BUDNAME FROM VIEW_BUDGET" & vbCrLf
        'sql &= " WHERE 1=1" & vbCrLf
        'sql &= " AND BUDID IN ('01','02','03')" & vbCrLf
        ''sql += " AND budid <='97'" & vbCrLf
        ''sql += " AND budid <>'04'" & vbCrLf
        'sql &= " ORDER BY BUDID" & vbCrLf
        'dt = DbAccess.GetDataTable(sql, objconn)
        'If dt.Rows.Count > 0 Then
        '    With BudID
        '        .DataSource = dt
        '        .DataTextField = "BudName"
        '        .DataValueField = "BudID"
        '        .DataBind()
        '    End With
        'End If

    End Sub

    '列印
    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click

        Dim s_OCIDStr As String = "" '勾選班級
        Dim s_TPlanID1 As String = "" '訓練計畫
        Dim s_DistID1 As String = "" '轄區參數
        Dim s_Identity1 As String = "" '身分別's_Identity1 = ""
        Dim sBudID As String = "" '預算別

        '報表要用的身分別參數
        s_Identity1 = ""
        For i As Integer = 1 To Me.Identity.Items.Count - 1
            If Me.Identity.Items(i).Selected Then
                If s_Identity1 <> "" Then s_Identity1 += ","
                s_Identity1 += Convert.ToString("" & Me.Identity.Items(i).Value & "")
            End If
        Next

        '報表要用的   '預算別
        sBudID = ""
        For i As Integer = 0 To Me.BudID.Items.Count - 1
            If Me.BudID.Items(i).Selected Then
                If sBudID <> "" Then sBudID += ","
                sBudID += Convert.ToString("" & Me.BudID.Items(i).Value & "")
            End If
        Next

        '報表要用的轄區參數
        s_DistID1 = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If s_DistID1 <> "" Then s_DistID1 += ","
                s_DistID1 += Convert.ToString("" & Me.DistID.Items(i).Value & "")
            End If
        Next


        '報表要用的 訓練計畫參數
        s_TPlanID1 = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected Then
                If s_TPlanID1 <> "" Then s_TPlanID1 += ","
                s_TPlanID1 += Convert.ToString("" & Me.TPlanID.Items(i).Value & "")
            End If
        Next

        ''勾選班級後會省略結訓日期的條件
        'If (RIDValue.Value <> "") OrElse (PlanID.Value <> "") Then
        '    FTDate1.Text = ""
        '    FTDate2.Text = ""
        '    STDate1.Text = ""
        '    STDate2.Text = ""
        '    Syear.SelectedIndex = -1
        'End If

        s_OCIDStr = ""
        'OCIDName = ""
        For Each item As ListItem In OCID.Items
            If item.Selected = True Then
                If item.Value = "%" Then '有選擇全選
                    s_OCIDStr = ""
                    'OCIDName = ""
                    For i As Integer = 1 To Me.OCID.Items.Count - 1
                        If s_OCIDStr <> "" Then s_OCIDStr += ","
                        s_OCIDStr += Convert.ToString(Me.OCID.Items(i).Value)
                    Next
                    Exit For
                Else
                    If s_OCIDStr <> "" Then s_OCIDStr += ","
                    s_OCIDStr += item.Value
                End If
            End If
        Next

        '勾選班級後會省略結訓日期的條件
        If s_OCIDStr <> "" Then
            '確保有效範圍
            If (RIDValue.Value <> "") OrElse (PlanID.Value <> "") Then
                FTDate1.Text = ""
                FTDate2.Text = ""

                STDate1.Text = ""
                STDate2.Text = ""
                Syear.SelectedIndex = -1
            End If
        End If

        Dim myValue As String = ""
        myValue = ""
        myValue += "&STTDate=" & Me.STDate1.Text '開訓起
        myValue += "&FTTDate=" & Me.STDate2.Text '開訓迄
        myValue += "&SFTDate=" & Me.FTDate1.Text '結訓起
        myValue += "&FFTDate=" & Me.FTDate2.Text '結訓迄
        If s_DistID1 <> "" Then myValue += "&DistID=" & s_DistID1
        If s_TPlanID1 <> "" Then myValue += "&TPlanID=" & s_TPlanID1
        If s_Identity1 <> "" Then myValue += "&Identity=" & s_Identity1
        If sBudID <> "" Then myValue += "&BudID=" & sBudID '預算別
        myValue += "&Years=" & Syear.SelectedValue '年度
        myValue += "&PlanID=" & PlanID.Value
        myValue += "&RID=" & RIDValue.Value
        If s_OCIDStr <> "" Then myValue += "&OCID=" & s_OCIDStr '勾選班級後會省略 年度、開、結訓日期的條件
        'myValue += "&DistName=" & DistName
        'myValue += "&IdentityName=" & IdentityName
        'myValue += "&TPlanName=" & TPlanName

        Select Case searcha_type1.SelectedValue
            Case "1"
                '統計資料'報表
                Dim xfilename As String = cst_printFN2B
                If Syear.SelectedValue = "" Then xfilename = cst_printFN2C
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, xfilename, myValue)
            Case "2"
                '明細資料'報表
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2A, myValue)
            Case Else
                Common.MessageBox(Me, "查詢方式未選擇，請確認！")
                Exit Sub
        End Select

    End Sub

    '查詢班級
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        PlanID.Value = TIMS.ClearSQM(PlanID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If PlanID.Value = "" Then Exit Sub
        If RIDValue.Value = "" Then Exit Sub
        msg.Text = ""

        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        'Dim dr As DataRow = Nothing
        sql = ""
        sql &= " SELECT OCID"
        sql &= " ,dbo.FN_GET_CLASSCNAME(CLASSCNAME,CYCLTYPE) CLASSCNAME"
        sql &= " FROM CLASS_CLASSINFO"
        sql &= " WHERE 1=1"
        sql &= " AND NOTOPEN='N'"
        sql &= " AND ISSUCCESS='Y'"
        sql &= " AND PlanID='" & PlanID.Value & "'"
        sql &= " AND RID='" & RIDValue.Value & "'"
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count = 0 Then
            msg.Text = "查無此機構底下的班級"
            'msg.Visible = True
            OCID.Style("display") = "none"
            Exit Sub
        End If

        With OCID.Items
            .Clear()
            .Add(New ListItem("全選", "%"))
            For Each dr As DataRow In dt.Rows
                .Add(New ListItem(dr("CLASSCNAME"), dr("OCID")))
            Next
        End With
        msg.Text = ""
        OCID.Style("display") = "inline"
        'msg.Visible = False

    End Sub

    '訓練機構...
    Private Sub Button2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.ServerClick
        Dim DistID1 As String = ""
        Dim TPlanID1 As String = ""
        Dim msg As String = ""
        Dim N As Integer = 0
        Dim N1 As Integer = 0

        DistID1 = ""
        N = 0   '預設 N =0 表示沒有勾選轄區選項
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then '假如有勾選
                N = N + 1  '計算轄區勾選選項的數目
                If N = 1 Then '如果是勾選一個選項
                    DistID1 = Convert.ToString(Me.DistID.Items(i).Value) '取得選項的值
                End If
                If N = 2 Then '如果轄區勾選選項的數目=2
                    'Common.MessageBox(Me, "只能選擇一個轄區")
                    msg += "只能選擇一個轄區!" & vbCrLf
                    DistID1 = ""
                    Exit For
                End If
            End If
        Next
        If N = 0 Then '如果轄區選項沒有選
            'Common.MessageBox(Me, "請選擇轄區")
            msg += "請選擇轄區!" & vbCrLf
        End If
        TPlanID1 = ""
        N1 = 0 '預設 N1 =0 表示沒有勾選計畫選項
        For j As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(j).Selected Then '假如有勾選
                N1 = N1 + 1 '計算計畫勾選選項的數目
                If N1 = 1 Then '如果是勾選一個選項
                    TPlanID1 = Convert.ToString(Me.TPlanID.Items(j).Value) '取得選項的值
                End If
                If N1 = 2 Then '如果計畫勾選選項的數目=2
                    'Common.MessageBox(Me, "只能選擇一個計畫")
                    msg += "只能選擇一個計畫!" & vbCrLf
                    TPlanID1 = ""
                    Exit For
                End If
            End If
        Next
        If N1 = 0 Then '如果計畫選項沒有選
            'Common.MessageBox(Me, "請選擇計畫")
            msg += "請選擇計畫!" & vbCrLf
        End If
        If msg <> "" Then
            Common.MessageBox(Me, msg)
        End If
        If DistID1 <> "" And TPlanID1 <> "" Then
            Dim strScript1 As String
            strScript1 = "<script language=""javascript"">" + vbCrLf
            strScript1 += "wopen('../../Common/MainOrg.aspx?DistID=' + '" & DistID1 & "' + '&TPlanID=' + '" & TPlanID1 & "'  + '&BtnName=Button3','查詢機構',400,400,1);"
            strScript1 += "</script>"
            Page.RegisterStartupScript("", strScript1)
        End If

    End Sub

    ''' <summary>
    ''' 查詢 (匯出Excel用) [SQL] 2020
    ''' </summary>
    ''' <param name="sType"></param>
    Sub Search4(ByVal sType As Integer)
        'sType 1:統計資料 2:明細資料
        Dim OCIDStr As String = ""
        Dim TPlanID1 As String = ""
        Dim DistID1 As String = ""
        Dim Identity1 As String = ""
        Dim sBudID As String = ""

        '身分別參數
        Identity1 = ""
        For i As Integer = 1 To Me.Identity.Items.Count - 1
            If Me.Identity.Items(i).Selected Then
                If Identity1 <> "" Then Identity1 += ","
                Identity1 += Convert.ToString("'" & Me.Identity.Items(i).Value & "'")
            End If
        Next

        '預算別
        sBudID = ""
        For i As Integer = 0 To Me.BudID.Items.Count - 1
            If Me.BudID.Items(i).Selected Then
                If sBudID <> "" Then sBudID += ","
                sBudID += Convert.ToString("'" & Me.BudID.Items(i).Value & "'")
            End If
        Next

        '轄區參數
        DistID1 = ""
        'DistName = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 += ","
                DistID1 += Convert.ToString("'" & Me.DistID.Items(i).Value & "'")
            End If
        Next

        '訓練計畫參數
        TPlanID1 = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected Then
                If TPlanID1 <> "" Then TPlanID1 += ","
                TPlanID1 += Convert.ToString("'" & Me.TPlanID.Items(i).Value & "'")
            End If
        Next

        OCIDStr = ""
        For Each item As ListItem In OCID.Items
            If item.Selected = True Then
                If item.Value = "%" Then '有選擇全選
                    OCIDStr = ""
                    For i As Integer = 1 To Me.OCID.Items.Count - 1
                        If OCIDStr <> "" Then OCIDStr += ","
                        OCIDStr += Convert.ToString(Me.OCID.Items(i).Value)
                    Next
                    Exit For
                Else
                    If OCIDStr <> "" Then OCIDStr += ","
                    OCIDStr += item.Value
                End If
            End If
        Next

        '勾選班級後會省略結訓日期的條件
        If OCIDStr <> "" Then
            If (RIDValue.Value <> "") OrElse (PlanID.Value <> "") Then
                FTDate1.Text = ""
                FTDate2.Text = ""

                STDate1.Text = ""
                STDate2.Text = ""
                Syear.SelectedIndex = -1
            End If
        End If

        Dim sql As String = ""
        'sType 1:統計資料 2:明細資料
        Select Case sType
            Case 1 'cst_2020_1
                'Dim sql As String = ""
                sql = "" & vbCrLf
                sql &= " WITH WKRT6 AS (SELECT SORT06,RTReasonID,Reason FROM Key_RejectTReason WHERE SORT06 IS NOT NULL )" & vbCrLf
                sql &= " ,WKRT3 AS (SELECT SORT3,RTReasonID,Reason FROM Key_RejectTReason WHERE SORT3 IS NOT NULL)" & vbCrLf
                sql &= " ,WGS1 AS (" & vbCrLf
                sql &= " select ip.years" & vbCrLf
                sql &= " ,ip.distid" & vbCrLf
                sql &= " ,cs.ocid" & vbCrLf
                sql &= " ,1 opencount" & vbCrLf
                sql &= " ,case when cs.Studstatus not in (2,3) then 1 END closecount" & vbCrLf
                'sql &= " /*04.患病或遇意外傷害/03.遇家庭等災變事故/07.自願、接受徵集入營者/31.工作異動/32.課程內容不符預期/01.缺課時數超過規定/99.其他*/" & vbCrLf
                'sql &= " /*13.參訓期間行為不檢情節重大/33.身分不符/99.其他*/" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='04' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x04" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='03' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x03" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='07' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x07" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='31' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x31" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='32' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x32" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='01' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x01" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2) and cs.RTReasonID ='99' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x98" & vbCrLf

                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='13' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x13" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='33' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x33" & vbCrLf
                sql &= " ,case when cs.Studstatus in (3) and cs.RTReasonID ='99' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x99" & vbCrLf

                sql &= " ,case when cs.Studstatus in (2,3) and (WKRT6.RTReasonID is not null OR WKRT3.RTReasonID is not null)" & vbCrLf
                sql &= " and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end xALL" & vbCrLf

                sql &= " ,case when cs.StudStatus IN (2,3) and cs.RejectDayIn14='Y' then 1 end sum_RDayIn14 /*遞補期內離訓人數*/" & vbCrLf
                sql &= " FROM class_classinfo cc WITH(NOLOCK)" & vbCrLf
                sql &= " JOIN plan_planinfo pp WITH(NOLOCK) on pp.planid =cc.planid and pp.comidno=cc.comidno and pp.seqno =cc.seqno" & vbCrLf
                sql &= " JOIN org_orginfo oo WITH(NOLOCK) on oo.comidno =cc.comidno" & vbCrLf
                sql &= " JOIN view_plan ip WITH(NOLOCK) on ip.planid=cc.planid" & vbCrLf
                sql &= " JOIN Class_StudentsOfClass cs WITH(NOLOCK) on cc.ocid =cs.ocid" & vbCrLf
                sql &= " LEFT JOIN WKRT3 ON WKRT3.RTReasonID=cs.RTReasonID" & vbCrLf
                sql &= " LEFT JOIN WKRT6 ON WKRT6.RTReasonID=cs.RTReasonID" & vbCrLf
                sql &= " where 1=1" & vbCrLf
                sql &= " and cc.NotOpen='N'" & vbCrLf
                sql &= " and cc.IsSuccess='Y'" & vbCrLf
                sql &= " and cc.FTDate < GETDATE()" & vbCrLf
                sql &= " and cs.MakeSOCID is null" & vbCrLf

                If Identity1 <> "" Then sql &= " and cs.MIdentityID IN (" & Identity1 & ")" & vbCrLf
                If sBudID <> "" Then sql &= " and cs.BudgetID IN (" & sBudID & ")" & vbCrLf
                If PlanID.Value <> "" Then sql &= " and cc.PlanID ='" & PlanID.Value & "'" & vbCrLf
                If RIDValue.Value <> "" Then sql &= " and cc.RID ='" & RIDValue.Value & "'" & vbCrLf
                If OCIDStr <> "" Then sql &= " and cc.OCID IN (" & OCIDStr & ")" & vbCrLf
                If Syear.SelectedValue <> "" Then sql &= "  and ip.years ='" & Syear.SelectedValue & "'" & vbCrLf
                If DistID1 <> "" Then sql &= " and ip.DistID IN (" & DistID1 & ")" & vbCrLf
                If TPlanID1 <> "" Then sql &= " and ip.TPlanID IN (" & TPlanID1 & ")" & vbCrLf
                If STDate1.Text <> "" Then sql &= " and cc.STDate>=" & TIMS.To_date(STDate1.Text) & vbCrLf
                If STDate2.Text <> "" Then sql &= " and cc.STDate<=" & TIMS.To_date(STDate2.Text) & vbCrLf
                If FTDate1.Text <> "" Then sql &= " and cc.FTDate>=" & TIMS.To_date(FTDate1.Text) & vbCrLf
                If FTDate2.Text <> "" Then sql &= " and cc.FTDate<=" & TIMS.To_date(FTDate2.Text) & vbCrLf

                sql &= " )" & vbCrLf
                sql &= " ,WGS2 AS (" & vbCrLf
                sql &= " select a.Years,a.DISTID" & vbCrLf
                sql &= " ,COUNT(a.opencount) opencount" & vbCrLf
                sql &= " ,COUNT(a.closecount) closecount" & vbCrLf

                sql &= " ,COUNT(a.x04) x04" & vbCrLf
                sql &= " ,COUNT(a.x03) x03" & vbCrLf
                sql &= " ,COUNT(a.x07) x07" & vbCrLf
                sql &= " ,COUNT(a.x31) x31" & vbCrLf
                sql &= " ,COUNT(a.x32) x32" & vbCrLf
                sql &= " ,COUNT(a.x01) x01" & vbCrLf
                sql &= " ,COUNT(a.x98) x98" & vbCrLf

                sql &= " ,COUNT(a.x13) x13" & vbCrLf
                sql &= " ,COUNT(a.x33) x33" & vbCrLf
                sql &= " ,COUNT(a.x99) x99" & vbCrLf
                sql &= " ,COUNT(a.xALL) xALL" & vbCrLf

                sql &= " ,COUNT(a.sum_RDayIn14) sum_RDayIn14" & vbCrLf
                sql &= " FROM WGS1 a" & vbCrLf
                sql &= " GROUP BY a.Years,a.DISTID" & vbCrLf
                sql &= " )" & vbCrLf

                sql &= " SELECT  A.YEARS" & vbCrLf
                sql &= " ,A.DISTID" & vbCrLf
                sql &= " ,K1.NAME DISTNAME" & vbCrLf
                sql &= " ,A.OPENCOUNT" & vbCrLf
                sql &= " ,a.CLOSECOUNT" & vbCrLf

                sql &= " ,a.X04" & vbCrLf
                sql &= " ,a.X03" & vbCrLf
                sql &= " ,a.X07" & vbCrLf
                sql &= " ,a.X31" & vbCrLf
                sql &= " ,a.X32" & vbCrLf
                sql &= " ,a.X01" & vbCrLf
                sql &= " ,a.X98" & vbCrLf

                sql &= " ,a.X13" & vbCrLf
                sql &= " ,a.X33" & vbCrLf
                sql &= " ,a.X99" & vbCrLf
                sql &= " ,a.XALL" & vbCrLf

                sql &= " ,a.SUM_RDAYIN14" & vbCrLf
                sql &= " FROM WGS2 a" & vbCrLf
                sql &= " JOIN ID_DISTRICT k1 on k1.distid=a.distid" & vbCrLf
                sql &= " ORDER BY a.YEARS,A.DISTID" & vbCrLf

            Case 2 'cst_2020_2
                'Dim sql As String = ""
                sql = "" & vbCrLf
                sql &= " WITH WKRT6 AS (SELECT SORT06,RTREASONID,REASON FROM Key_RejectTReason WHERE SORT06 IS NOT NULL )" & vbCrLf
                sql &= " ,WKRT3 AS (SELECT SORT3,RTREASONID,REASON FROM Key_RejectTReason WHERE SORT3 IS NOT NULL )" & vbCrLf
                sql &= " ,WGS1 AS (" & vbCrLf
                sql &= " select ip.YEARS" & vbCrLf
                sql &= " ,ip.distid" & vbCrLf
                sql &= " ,cs.ocid" & vbCrLf
                sql &= " ,1 opencount" & vbCrLf
                sql &= " ,case when cs.Studstatus not in (2,3) then 1 END closecount" & vbCrLf
                'sql &= " /*04.患病或遇意外傷害/03.遇家庭等災變事故/07.自願、接受徵集入營者/31.工作異動/32.課程內容不符預期/01.缺課時數超過規定/99.其他*/" & vbCrLf
                'sql &= " /*13.參訓期間行為不檢情節重大/33.身分不符/99.其他*/" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='04' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x04" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='03' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x03" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='07' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x07" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='31' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x31" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='32' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x32" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='01' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x01" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2) and cs.RTReasonID ='99' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x98" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='13' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x13" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='33' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x33" & vbCrLf
                sql &= " ,case when cs.Studstatus in (3) and cs.RTReasonID ='99' and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end x99" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and (WKRT6.RTReasonID is not null OR WKRT3.RTReasonID is not null)" & vbCrLf
                sql &= " and ISNULL(cs.RejectDayIn14,'N')!='Y' then 1 end xALL" & vbCrLf
                sql &= " ,case when cs.StudStatus IN (2,3) and cs.RejectDayIn14='Y' then 1 end sum_RDayIn14 /*遞補期內離訓人數*/" & vbCrLf
                sql &= " FROM dbo.CLASS_CLASSINFO cc" & vbCrLf
                sql &= " JOIN dbo.PLAN_PLANINFO pp on pp.planid =cc.planid and pp.comidno=cc.comidno and pp.seqno =cc.seqno" & vbCrLf
                sql &= " JOIN dbo.ORG_ORGINFO oo on oo.comidno =cc.comidno" & vbCrLf
                sql &= " JOIN dbo.VIEW_PLAN ip on ip.planid=cc.planid" & vbCrLf
                sql &= " JOIN dbo.CLASS_STUDENTSOFCLASS cs on cc.ocid =cs.ocid" & vbCrLf
                sql &= " LEFT JOIN WKRT3 ON WKRT3.RTReasonID=cs.RTReasonID" & vbCrLf
                sql &= " LEFT JOIN WKRT6 ON WKRT6.RTReasonID=cs.RTReasonID" & vbCrLf
                sql &= " where 1=1" & vbCrLf
                sql &= " and cc.NotOpen='N'" & vbCrLf
                sql &= " and cc.IsSuccess='Y'" & vbCrLf
                sql &= " and cc.FTDate < GETDATE()" & vbCrLf
                sql &= " and cs.MakeSOCID IS NULL" & vbCrLf

                If Identity1 <> "" Then sql &= " and cs.MIdentityID IN (" & Identity1 & ")" & vbCrLf
                If sBudID <> "" Then sql &= " and cs.BudgetID IN (" & sBudID & ")" & vbCrLf
                If PlanID.Value <> "" Then sql &= " and cc.PlanID ='" & PlanID.Value & "'" & vbCrLf
                If RIDValue.Value <> "" Then sql &= " and cc.RID ='" & RIDValue.Value & "'" & vbCrLf
                If OCIDStr <> "" Then sql &= " and cc.OCID IN (" & OCIDStr & ")" & vbCrLf
                If Syear.SelectedValue <> "" Then sql &= "  and ip.years ='" & Syear.SelectedValue & "'" & vbCrLf
                If DistID1 <> "" Then sql &= " and ip.DistID IN (" & DistID1 & ")" & vbCrLf
                If TPlanID1 <> "" Then sql &= " and ip.TPlanID IN (" & TPlanID1 & ")" & vbCrLf
                If STDate1.Text <> "" Then sql &= " and cc.STDate>=" & TIMS.To_date(STDate1.Text) & vbCrLf
                If STDate2.Text <> "" Then sql &= " and cc.STDate<=" & TIMS.To_date(STDate2.Text) & vbCrLf
                If FTDate1.Text <> "" Then sql &= " and cc.FTDate>=" & TIMS.To_date(FTDate1.Text) & vbCrLf
                If FTDate2.Text <> "" Then sql &= " and cc.FTDate<=" & TIMS.To_date(FTDate2.Text) & vbCrLf

                'sql &= " AND ip.years ='2019' AND ip.TPLANID ='06'" & vbCrLf
                'sql &= " /* AND ip.years ='2019' AND ip.TPLANID ='06' */" & vbCrLf
                sql &= " )" & vbCrLf
                sql &= " ,WGS2 AS (" & vbCrLf
                sql &= " select a.OCID" & vbCrLf
                sql &= " ,COUNT(a.opencount) opencount" & vbCrLf
                sql &= " ,COUNT(a.closecount) closecount" & vbCrLf
                sql &= " ,COUNT(a.x04) x04" & vbCrLf
                sql &= " ,COUNT(a.x03) x03" & vbCrLf
                sql &= " ,COUNT(a.x07) x07" & vbCrLf
                sql &= " ,COUNT(a.x31) x31" & vbCrLf
                sql &= " ,COUNT(a.x32) x32" & vbCrLf
                sql &= " ,COUNT(a.x01) x01" & vbCrLf
                sql &= " ,COUNT(a.x98) x98" & vbCrLf
                sql &= " ,COUNT(a.x13) x13" & vbCrLf
                sql &= " ,COUNT(a.x33) x33" & vbCrLf
                sql &= " ,COUNT(a.x99) x99" & vbCrLf
                sql &= " ,COUNT(a.xALL) xALL" & vbCrLf
                sql &= " ,COUNT(a.sum_RDayIn14) sum_RDayIn14" & vbCrLf
                sql &= " FROM WGS1 a" & vbCrLf
                sql &= " GROUP BY a.OCID" & vbCrLf
                sql &= " )" & vbCrLf

                sql &= " SELECT cc.YEARS" & vbCrLf
                sql &= " ,a.OCID" & vbCrLf
                sql &= " ,cc.DISTID" & vbCrLf
                sql &= " ,cc.DISTNAME" & vbCrLf
                sql &= " ,cc.TPLANID" & vbCrLf
                sql &= " ,cc.PLANNAME" & vbCrLf
                sql &= " ,cc.ORGID" & vbCrLf
                sql &= " ,cc.ORGNAME" & vbCrLf
                sql &= " ,cc.CLASSCNAME2" & vbCrLf
                sql &= " ,CONVERT(VARCHAR,cc.stdate,111) STDATE" & vbCrLf
                sql &= " ,CONVERT(VARCHAR,cc.ftdate,111) FTDATE" & vbCrLf
                sql &= " ,a.OPENCOUNT" & vbCrLf
                sql &= " ,a.CLOSECOUNT" & vbCrLf
                sql &= " ,a.X04" & vbCrLf
                sql &= " ,a.X03" & vbCrLf
                sql &= " ,a.X07" & vbCrLf
                sql &= " ,a.X31" & vbCrLf
                sql &= " ,a.X32" & vbCrLf
                sql &= " ,a.X01" & vbCrLf
                sql &= " ,a.X98" & vbCrLf
                sql &= " ,a.X13" & vbCrLf
                sql &= " ,a.X33" & vbCrLf
                sql &= " ,a.X99" & vbCrLf
                sql &= " ,a.XALL" & vbCrLf
                sql &= " ,a.SUM_RDAYIN14" & vbCrLf
                sql &= " FROM WGS2 a" & vbCrLf
                sql &= " JOIN VIEW2 cc on cc.ocid =a.ocid" & vbCrLf
                sql &= " ORDER BY cc.YEARS,CC.DISTID,CC.ORGNAME,CC.CLASSCNAME2" & vbCrLf

        End Select

        Call TIMS.OpenDbConn(objconn)
        Dim oCmd As New SqlCommand(sql, objconn)
        Dim dt1 As New DataTable
        With oCmd
            .Parameters.Clear()
            dt1.Load(.ExecuteReader())
        End With

        Const cst_dcad_年度 As String = "年度"
        Const cst_dcad_轄區 As String = "轄區"

        Const cst_dcad_訓練計畫 As String = "訓練計畫"
        Const cst_dcad_培訓單位 As String = "培訓單位"
        Const cst_dcad_班別 As String = "班別"
        Const cst_dcad_開訓日期 As String = "開訓日期"
        Const cst_dcad_結訓日期 As String = "結訓日期"

        Const cst_dcad_開訓人數 As String = "開訓人數"
        Const cst_dcad_結訓人數 As String = "結訓人數"

        Const cst_dcad_患病或遇意外傷害 As String = "患病或遇意外傷害" 'X04
        Const cst_dcad_遇家庭等災變事故 As String = "遇家庭等災變事故" 'X03
        Const cst_dcad_自願接受徵集入營者 As String = "自願、接受徵集入營者" 'X07
        Const cst_dcad_工作異動 As String = "工作異動" 'X31
        Const cst_dcad_課程內容不符預期 As String = "課程內容不符預期" 'X32
        Const cst_dcad_缺課時數超過規定 As String = "缺課時數超過規定" 'X01
        Const cst_dcad_其他離訓 As String = "其他(離訓)" 'X98

        Const cst_dcad_參訓期間行為不檢情節重大 As String = "參訓期間行為不檢情節重大" 'X13
        Const cst_dcad_身分不符 As String = "身分不符" 'X33
        Const cst_dcad_其他退訓 As String = "其他(退訓)" 'X99

        Const cst_dcad_合計 As String = "合計"

        'https://msdn.microsoft.com/zh-tw/library/system.data.datacolumn.datatype(v=vs.110).aspx
        Dim dt As New DataTable
        dt.Columns.Add(New DataColumn(cst_dcad_年度))
        dt.Columns.Add(New DataColumn(cst_dcad_轄區))
        Select Case sType
            Case 1
            Case Else
                dt.Columns.Add(New DataColumn(cst_dcad_訓練計畫))
                dt.Columns.Add(New DataColumn(cst_dcad_培訓單位))
                dt.Columns.Add(New DataColumn(cst_dcad_班別))
                dt.Columns.Add(New DataColumn(cst_dcad_開訓日期))
                dt.Columns.Add(New DataColumn(cst_dcad_結訓日期))

        End Select
        dt.Columns.Add(New DataColumn(cst_dcad_開訓人數, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_結訓人數, System.Type.GetType("System.Int32")))

        dt.Columns.Add(New DataColumn(cst_dcad_患病或遇意外傷害, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_遇家庭等災變事故, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_自願接受徵集入營者, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_工作異動, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_課程內容不符預期, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_缺課時數超過規定, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_其他離訓, System.Type.GetType("System.Int32")))

        dt.Columns.Add(New DataColumn(cst_dcad_參訓期間行為不檢情節重大, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_身分不符, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_其他退訓, System.Type.GetType("System.Int32")))

        dt.Columns.Add(New DataColumn(cst_dcad_合計, System.Type.GetType("System.Int32")))

        '合計
        If dt1.Rows.Count > 0 Then
            Dim int_開訓人數 As Integer = 0
            Dim int_結訓人數 As Integer = 0

            Dim int_X04 As Integer = 0
            Dim int_X03 As Integer = 0
            Dim int_X07 As Integer = 0
            Dim int_X31 As Integer = 0
            Dim int_X32 As Integer = 0
            Dim int_X01 As Integer = 0
            Dim int_X98 As Integer = 0

            Dim int_X13 As Integer = 0
            Dim int_X33 As Integer = 0
            Dim int_X99 As Integer = 0
            Dim int_合計 As Integer = 0

            Dim dr As DataRow = Nothing
            For Each dr1 As DataRow In dt1.Rows
                int_開訓人數 += dr1("opencount")
                int_結訓人數 += dr1("closecount")

                int_X04 += dr1("X04")
                int_X03 += dr1("X03")
                int_X07 += dr1("X07")
                int_X31 += dr1("X31")
                int_X32 += dr1("X32")
                int_X01 += dr1("X01")
                int_X98 += dr1("X98")

                int_X13 += dr1("X13")
                int_X33 += dr1("X33")
                int_X99 += dr1("X99")
                int_合計 += dr1("xALL")

                dr = dt.NewRow()
                dr(cst_dcad_年度) = dr1("Years")
                dr(cst_dcad_轄區) = dr1("distname")

                Select Case sType
                    Case 1
                        dr(cst_dcad_年度) = dr1("Years")
                        dr(cst_dcad_轄區) = dr1("distname")

                    Case Else
                        dr(cst_dcad_年度) = dr1("Years")
                        dr(cst_dcad_轄區) = dr1("distname")
                        dr(cst_dcad_訓練計畫) = dr1("planname")
                        dr(cst_dcad_培訓單位) = dr1("orgname")
                        dr(cst_dcad_班別) = dr1("classcname2")
                        dr(cst_dcad_開訓日期) = dr1("stdate")
                        dr(cst_dcad_結訓日期) = dr1("ftdate")

                End Select

                dr(cst_dcad_開訓人數) = dr1("opencount")
                dr(cst_dcad_結訓人數) = dr1("closecount")

                dr(cst_dcad_患病或遇意外傷害) = dr1("X04")
                dr(cst_dcad_遇家庭等災變事故) = dr1("X03")
                dr(cst_dcad_自願接受徵集入營者) = dr1("X07")
                dr(cst_dcad_工作異動) = dr1("X31")
                dr(cst_dcad_課程內容不符預期) = dr1("X32")
                dr(cst_dcad_缺課時數超過規定) = dr1("X01")
                dr(cst_dcad_其他離訓) = dr1("X98")

                dr(cst_dcad_參訓期間行為不檢情節重大) = dr1("X13")
                dr(cst_dcad_身分不符) = dr1("X33")
                dr(cst_dcad_其他退訓) = dr1("X99")
                dr(cst_dcad_合計) = dr1("xALL")

                dt.Rows.Add(dr)
            Next

            '最後1行加總
            'sType 1:統計資料 2:明細資料
            dr = dt.NewRow()
            Select Case sType
                Case 1
                    dr(cst_dcad_年度) = cst_dcad_合計
                    dr(cst_dcad_轄區) = " "
                Case 2
                    dr(cst_dcad_年度) = cst_dcad_合計
                    dr(cst_dcad_轄區) = " "
                    dr(cst_dcad_訓練計畫) = " "
                    dr(cst_dcad_培訓單位) = " "
                    dr(cst_dcad_班別) = " "
                    dr(cst_dcad_開訓日期) = " "
                    dr(cst_dcad_結訓日期) = " "
            End Select

            dr(cst_dcad_開訓人數) = int_開訓人數
            dr(cst_dcad_結訓人數) = int_結訓人數

            dr(cst_dcad_患病或遇意外傷害) = int_X04
            dr(cst_dcad_遇家庭等災變事故) = int_X03
            dr(cst_dcad_自願接受徵集入營者) = int_X07
            dr(cst_dcad_工作異動) = int_X31
            dr(cst_dcad_課程內容不符預期) = int_X32
            dr(cst_dcad_缺課時數超過規定) = int_X01
            dr(cst_dcad_其他離訓) = int_X98

            dr(cst_dcad_參訓期間行為不檢情節重大) = int_X13
            dr(cst_dcad_身分不符) = int_X33
            dr(cst_dcad_其他退訓) = int_X99
            dr(cst_dcad_合計) = int_合計

            dt.Rows.Add(dr)
        End If
        'dt.AcceptChanges()

        ExportMsg.Text = "查無資料"
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            ExportMsg.Text = ""

            With DataGrid1
                .Visible = True
                .DataSource = dt
                .DataBind()
            End With

        End If
    End Sub

#Region "OLD_2020"

    '查詢 (匯出Excel用) [SQL] 2017
    Sub Search3(ByVal sType As Integer)
        'sType 1:統計資料 2:明細資料
        Dim OCIDStr As String = ""
        Dim TPlanID1 As String = ""
        Dim DistID1 As String = ""
        Dim Identity1 As String = ""
        Dim sBudID As String = ""

        '身分別參數
        Identity1 = ""
        For i As Integer = 1 To Me.Identity.Items.Count - 1
            If Me.Identity.Items(i).Selected Then
                If Identity1 <> "" Then Identity1 += ","
                Identity1 += Convert.ToString("'" & Me.Identity.Items(i).Value & "'")
            End If
        Next

        '預算別
        sBudID = ""
        For i As Integer = 0 To Me.BudID.Items.Count - 1
            If Me.BudID.Items(i).Selected Then
                If sBudID <> "" Then sBudID += ","
                sBudID += Convert.ToString("'" & Me.BudID.Items(i).Value & "'")
            End If
        Next

        '轄區參數
        DistID1 = ""
        'DistName = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 += ","
                DistID1 += Convert.ToString("'" & Me.DistID.Items(i).Value & "'")
            End If
        Next

        '訓練計畫參數
        TPlanID1 = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected Then
                If TPlanID1 <> "" Then TPlanID1 += ","
                TPlanID1 += Convert.ToString("'" & Me.TPlanID.Items(i).Value & "'")
            End If
        Next

        OCIDStr = ""
        For Each item As ListItem In OCID.Items
            If item.Selected = True Then
                If item.Value = "%" Then '有選擇全選
                    OCIDStr = ""
                    For i As Integer = 1 To Me.OCID.Items.Count - 1
                        If OCIDStr <> "" Then OCIDStr += ","
                        OCIDStr += Convert.ToString(Me.OCID.Items(i).Value)
                    Next
                    Exit For
                Else
                    If OCIDStr <> "" Then OCIDStr += ","
                    OCIDStr += item.Value
                End If
            End If
        Next

        '勾選班級後會省略結訓日期的條件
        If OCIDStr <> "" Then
            If (RIDValue.Value <> "") OrElse (PlanID.Value <> "") Then
                FTDate1.Text = ""
                FTDate2.Text = ""

                STDate1.Text = ""
                STDate2.Text = ""
                Syear.SelectedIndex = -1
            End If
        End If

        Dim sql As String = ""
        'sType 1:統計資料 2:明細資料
        Select Case sType
            Case 1
                'Dim sql As String = ""
                sql = "" & vbCrLf
                sql &= " WITH WKRT2 AS (SELECT dbo.NVL(SORT22,SORT2) SORT2,RTReasonID,Reason FROM Key_RejectTReason WHERE dbo.NVL(SORT22,SORT2) IS NOT NULL)" & vbCrLf
                sql &= " ,WKRT3 AS (SELECT SORT3,RTReasonID,Reason FROM Key_RejectTReason WHERE SORT3 IS NOT NULL )" & vbCrLf
                sql &= " ,WKRT6 AS (SELECT SORT06,RTReasonID,Reason FROM Key_RejectTReason WHERE SORT06 IS NOT NULL )" & vbCrLf
                'sql &= " ,WTD1 AS (select WM_CONCAT(CONVERT(varchar, Name)) TDistName FROM ID_District where 1=1)" & vbCrLf
                'sql &= " ,WTP1 AS (select WM_CONCAT(CONVERT(varchar, planname)) TPlanName FROM key_plan where 1=1)" & vbCrLf
                'sql &= " ,WID1 AS (select WM_CONCAT(CONVERT(varchar, Name)) IDName FROM Key_Identity where 1=1)" & vbCrLf
                sql &= " ,WGS1 AS (" & vbCrLf
                sql &= " select ip.years" & vbCrLf
                sql &= " ,ip.distid" & vbCrLf
                sql &= " ,cs.ocid" & vbCrLf
                sql &= " ,1 opencount" & vbCrLf
                sql &= " ,case when cs.Studstatus not in (2,3) then 1 END closecount" & vbCrLf
                sql &= " ,case when cs.Studstatus not in (2,3) and sg3.socid is not null then 1 end jobcount" & vbCrLf
                sql &= " ,case when cs.Studstatus not in (2,3) AND sg3.JOBRELATE='Y' then 1 end  JobrelxCount" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='04' then 1 end x04" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='03' then 1 end x03" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='07' then 1 end x07" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='31' then 1 end x31" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='32' then 1 end x32" & vbCrLf

                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='02' then 1 end x02" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='20' then 1 end x20" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='21' then 1 end x21" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='22' then 1 end x22" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID IN ('98','23') then 1 end x98" & vbCrLf
                'sql &= " /*SureItem: 1:雇主切結 2:學員切結 3:勞保勾稽(null)" & vbCrLf
                'sql &= " 提前就業-勞保勾稽:3 提前就業-學員切結:2 提前就業-雇主切結:1*/" & vbCrLf
                sql &= " ,CASE WHEN cs.WkAheadOfSch='Y' and cs.StudStatus in (2,3) AND sg9.mode_=1 and dbo.NVL(sg9.SureItem,'3')='3' THEN 1 END j9x3" & vbCrLf
                sql &= " ,CASE WHEN cs.WkAheadOfSch='Y' and cs.StudStatus in (2,3) AND sg9.mode_=2 and dbo.NVL(sg9.SureItem,'3')='2' THEN 1 END j9x2" & vbCrLf
                sql &= " ,CASE WHEN cs.WkAheadOfSch='Y' and cs.StudStatus in (2,3) AND sg9.mode_=2 and dbo.NVL(sg9.SureItem,'3')='1' THEN 1 END j9x1" & vbCrLf
                'sql &= " " & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='01' then 1 end x01" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='13' then 1 end x13" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='14' then 1 end x14" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='99' then 1 end x99" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3)" & vbCrLf
                sql &= "    and (WKRT2.RTReasonID is not null or WKRT3.RTReasonID is not null or WKRT6.RTReasonID is not null)" & vbCrLf
                sql &= "    and cs.RTReasonID !='02' then 1 end xALL" & vbCrLf

                'sql &= " /*「公法救助」*/" & vbCrLf
                sql &= " ,CASE WHEN cs.StudStatus not in (2,3) and sg3.PUBLICRESCUE='Y' AND sg3.SOCID IS NOT NULL THEN 1 END sum_PUBLICRESCUE" & vbCrLf
                'sql &= " /*提前就業「公法救助」*/" & vbCrLf
                sql &= " ,CASE WHEN cs.StudStatus in (2,3) and sg9.PUBLICRESCUE='Y' AND sg9.SOCID IS NOT NULL THEN 1 END sum_PUBLICRESCUE9" & vbCrLf
                sql &= " ,case when cs.StudStatus NOT IN (2,3) and sg3.JOBRELATE='Y' then 1 end sum_jobrelx /*就業關聯性*/" & vbCrLf
                sql &= " FROM class_classinfo cc" & vbCrLf
                sql &= " JOIN plan_planinfo pp on pp.planid =cc.planid and pp.comidno=cc.comidno and pp.seqno =cc.seqno" & vbCrLf
                sql &= " JOIN org_orginfo oo on oo.comidno =cc.comidno" & vbCrLf
                sql &= " JOIN view_plan ip on ip.planid=cc.planid" & vbCrLf
                sql &= " JOIN Class_StudentsOfClass cs on cc.ocid =cs.ocid" & vbCrLf
                sql &= " LEFT JOIN WKRT2 ON WKRT2.RTReasonID=cs.RTReasonID" & vbCrLf
                sql &= " LEFT JOIN WKRT3 ON WKRT3.RTReasonID=cs.RTReasonID" & vbCrLf
                sql &= " LEFT JOIN WKRT6 ON WKRT6.RTReasonID=cs.RTReasonID" & vbCrLf

                sql &= " LEFT JOIN Stud_GetJobState3 sg3 on sg3.socid =cs.socid and sg3.CPoint=1 and sg3.IsGetJob=1" & vbCrLf
                sql &= " LEFT JOIN Stud_GetJobState3 sg9 on sg9.socid =cs.socid and sg9.CPoint= 9" & vbCrLf
                sql &= " where 1=1" & vbCrLf
                sql &= " and cc.NotOpen='N'" & vbCrLf
                sql &= " and cc.IsSuccess='Y'" & vbCrLf
                sql &= " and cc.FTDate < getdate()" & vbCrLf
                sql &= " and cs.MakeSOCID is null" & vbCrLf
                'sql &= " AND ip.years ='2016'" & vbCrLf
                'sql &= " AND ip.TPLANID ='02'" & vbCrLf
                If Identity1 <> "" Then
                    sql &= " and cs.MIdentityID IN (" & Identity1 & ")" & vbCrLf
                End If
                If sBudID <> "" Then
                    sql &= " and cs.BudgetID IN (" & sBudID & ")" & vbCrLf
                End If
                If PlanID.Value <> "" Then
                    sql &= "  and cc.PlanID ='" & PlanID.Value & "'" & vbCrLf
                End If
                If RIDValue.Value <> "" Then
                    sql &= "  and cc.RID ='" & RIDValue.Value & "'" & vbCrLf
                End If
                If OCIDStr <> "" Then
                    sql &= "  and cc.OCID IN (" & OCIDStr & ")" & vbCrLf
                End If
                If Syear.SelectedValue <> "" Then
                    sql &= "  and ip.years ='" & Syear.SelectedValue & "'" & vbCrLf
                End If
                If DistID1 <> "" Then
                    sql &= "  and ip.DistID IN (" & DistID1 & ")" & vbCrLf
                End If
                If TPlanID1 <> "" Then
                    sql &= "  and ip.TPlanID IN (" & TPlanID1 & ")" & vbCrLf
                End If
                If STDate1.Text <> "" Then
                    sql &= "  and cc.STDate>=" & TIMS.To_date(STDate1.Text) & vbCrLf
                End If
                If STDate2.Text <> "" Then
                    sql &= "  and cc.STDate<=" & TIMS.To_date(STDate2.Text) & vbCrLf
                End If
                If FTDate1.Text <> "" Then
                    sql &= "  and cc.FTDate>=" & TIMS.To_date(FTDate1.Text) & vbCrLf
                End If
                If FTDate2.Text <> "" Then
                    sql &= "  and cc.FTDate<=" & TIMS.To_date(FTDate2.Text) & vbCrLf
                End If

                sql &= " )" & vbCrLf
                sql &= " ,WGS2 AS (" & vbCrLf
                sql &= " select a.Years,a.DISTID" & vbCrLf
                sql &= " ,COUNT(a.opencount) opencount" & vbCrLf
                sql &= " ,COUNT(a.closecount) closecount" & vbCrLf
                sql &= " ,COUNT(a.jobcount) jobcount" & vbCrLf
                sql &= " ,COUNT(a.JobrelxCount) JobrelxCount" & vbCrLf
                sql &= " ,COUNT(a.x04) x04" & vbCrLf
                sql &= " ,COUNT(a.x03) x03" & vbCrLf
                sql &= " ,COUNT(a.x07) x07" & vbCrLf
                sql &= " ,COUNT(a.x31) x31" & vbCrLf
                sql &= " ,COUNT(a.x32) x32" & vbCrLf

                sql &= " ,COUNT(a.x02) x02" & vbCrLf
                sql &= " ,COUNT(a.x20) x20" & vbCrLf
                sql &= " ,COUNT(a.x21) x21" & vbCrLf
                sql &= " ,COUNT(a.x22) x22" & vbCrLf
                sql &= " ,COUNT(a.x98) x98" & vbCrLf
                sql &= " ,COUNT(a.j9x3) j9x3" & vbCrLf
                sql &= " ,COUNT(a.j9x2) j9x2" & vbCrLf
                sql &= " ,COUNT(a.j9x1) j9x1" & vbCrLf
                'sql &= " " & vbCrLf
                sql &= " ,COUNT(a.x01) x01" & vbCrLf
                sql &= " ,COUNT(a.x13) x13" & vbCrLf
                sql &= " ,COUNT(a.x14) x14" & vbCrLf
                sql &= " ,COUNT(a.x99) x99" & vbCrLf
                sql &= " ,COUNT(a.xALL) xALL" & vbCrLf
                sql &= " ,COUNT(a.sum_PUBLICRESCUE) sum_PUBLICRESCUE" & vbCrLf
                sql &= " ,COUNT(a.sum_PUBLICRESCUE9) sum_PUBLICRESCUE9" & vbCrLf
                sql &= " ,COUNT(a.sum_jobrelx) sum_jobrelx" & vbCrLf
                sql &= " FROM WGS1 a" & vbCrLf
                sql &= " GROUP BY a.Years,a.DISTID" & vbCrLf
                sql &= " )" & vbCrLf
                'sql &= " SELECT WTD1.TDistName" & vbCrLf
                'sql &= " ,WTP1.TPlanName" & vbCrLf
                'sql &= " ,WID1.IDName" & vbCrLf
                sql &= " SELECT a.years" & vbCrLf
                sql &= " ,a.distid" & vbCrLf
                sql &= " ,k1.name distname" & vbCrLf
                sql &= " ,a.opencount" & vbCrLf
                sql &= " ,a.closecount" & vbCrLf
                sql &= " ,a.jobcount" & vbCrLf
                sql &= " ,a.JobrelxCount" & vbCrLf
                sql &= " ,a.x04" & vbCrLf
                sql &= " ,a.x03" & vbCrLf
                sql &= " ,a.x07" & vbCrLf
                sql &= " ,a.x31" & vbCrLf
                sql &= " ,a.x32" & vbCrLf

                sql &= " ,a.x02" & vbCrLf
                sql &= " ,a.x20" & vbCrLf
                sql &= " ,a.x21" & vbCrLf
                sql &= " ,a.x22" & vbCrLf
                sql &= " ,a.x98" & vbCrLf
                sql &= " ,a.j9x3" & vbCrLf
                sql &= " ,a.j9x2" & vbCrLf
                sql &= " ,a.j9x1" & vbCrLf
                'sql &= " " & vbCrLf
                sql &= " ,a.x01" & vbCrLf
                sql &= " ,a.x13" & vbCrLf
                sql &= " ,a.x14" & vbCrLf
                sql &= " ,a.x99" & vbCrLf
                sql &= " ,a.xALL" & vbCrLf
                sql &= " ,a.sum_PUBLICRESCUE" & vbCrLf
                sql &= " ,a.sum_PUBLICRESCUE9" & vbCrLf
                sql &= " ,a.sum_jobrelx" & vbCrLf
                sql &= " FROM WGS2 a" & vbCrLf
                sql &= " JOIN ID_DISTRICT k1 on k1.distid=a.distid" & vbCrLf
                'sql &= " CROSS JOIN WTD1" & vbCrLf
                'sql &= " CROSS JOIN WTP1" & vbCrLf
                'sql &= " CROSS JOIN WID1" & vbCrLf
                sql &= " ORDER BY a.DISTID" & vbCrLf

            Case 2
                'Dim sql As String = ""
                sql = "" & vbCrLf
                sql &= " WITH WKRT2 AS (SELECT dbo.NVL(SORT22,SORT2) SORT2,RTReasonID,Reason FROM Key_RejectTReason WHERE dbo.NVL(SORT22,SORT2) IS NOT NULL)" & vbCrLf
                sql &= " ,WKRT3 AS (SELECT SORT3,RTReasonID,Reason FROM Key_RejectTReason WHERE SORT3 IS NOT NULL )" & vbCrLf
                sql &= " ,WKRT6 AS (SELECT SORT06,RTReasonID,Reason FROM Key_RejectTReason WHERE SORT06 IS NOT NULL )" & vbCrLf
                'sql &= " ,WTD1 AS (select WM_CONCAT(CONVERT(varchar, Name)) TDistName FROM ID_District where 1=1)" & vbCrLf
                'sql &= " ,WTP1 AS (select WM_CONCAT(CONVERT(varchar, planname)) TPlanName FROM key_plan where 1=1)" & vbCrLf
                'sql &= " ,WID1 AS (select WM_CONCAT(CONVERT(varchar, Name)) IDName FROM Key_Identity where 1=1)" & vbCrLf
                sql &= " ,WGS1 AS (" & vbCrLf
                sql &= " select ip.years" & vbCrLf
                sql &= " ,ip.distid" & vbCrLf
                sql &= " ,cs.ocid" & vbCrLf
                sql &= " ,1 opencount" & vbCrLf
                sql &= " ,case when cs.Studstatus not in (2,3) then 1 END closecount" & vbCrLf
                sql &= " ,case when cs.Studstatus not in (2,3) and sg3.socid is not null then 1 end jobcount" & vbCrLf
                sql &= " ,case when cs.Studstatus not in (2,3) AND sg3.JOBRELATE='Y' then 1 end  JobrelxCount" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='04' then 1 end x04" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='03' then 1 end x03" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='07' then 1 end x07" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='31' then 1 end x31" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='32' then 1 end x32" & vbCrLf

                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='02' then 1 end x02" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='20' then 1 end x20" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='21' then 1 end x21" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='22' then 1 end x22" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID IN ('98','23') then 1 end x98" & vbCrLf
                'sql &= " /*SureItem: 1:雇主切結 2:學員切結 3:勞保勾稽(null)" & vbCrLf
                'sql &= " 提前就業-勞保勾稽:3 提前就業-學員切結:2 提前就業-雇主切結:1*/" & vbCrLf
                sql &= " ,CASE WHEN cs.WkAheadOfSch='Y' and cs.StudStatus in (2,3) AND sg9.mode_=1 and dbo.NVL(sg9.SureItem,'3')='3' THEN 1 END j9x3" & vbCrLf
                sql &= " ,CASE WHEN cs.WkAheadOfSch='Y' and cs.StudStatus in (2,3) AND sg9.mode_=2 and dbo.NVL(sg9.SureItem,'3')='2' THEN 1 END j9x2" & vbCrLf
                sql &= " ,CASE WHEN cs.WkAheadOfSch='Y' and cs.StudStatus in (2,3) AND sg9.mode_=2 and dbo.NVL(sg9.SureItem,'3')='1' THEN 1 END j9x1" & vbCrLf
                sql &= " " & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='01' then 1 end x01" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='13' then 1 end x13" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='14' then 1 end x14" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3) and cs.RTReasonID ='99' then 1 end x99" & vbCrLf
                sql &= " ,case when cs.Studstatus in (2,3)" & vbCrLf
                sql &= "    and (WKRT2.RTReasonID is not null or WKRT3.RTReasonID is not null or WKRT6.RTReasonID is not null)" & vbCrLf
                sql &= "    and cs.RTReasonID !='02' then 1 end xALL" & vbCrLf
                'sql &= " /*「公法救助」*/" & vbCrLf
                sql &= " ,CASE WHEN cs.StudStatus not in (2,3) and sg3.PUBLICRESCUE='Y' AND sg3.SOCID IS NOT NULL THEN 1 END sum_PUBLICRESCUE" & vbCrLf
                'sql &= " /*提前就業「公法救助」*/" & vbCrLf
                sql &= " ,CASE WHEN cs.StudStatus in (2,3) and sg9.PUBLICRESCUE='Y' AND sg9.SOCID IS NOT NULL THEN 1 END sum_PUBLICRESCUE9" & vbCrLf
                sql &= " ,case when cs.StudStatus NOT IN (2,3) and sg3.JOBRELATE='Y' then 1 end sum_jobrelx /*就業關聯性*/" & vbCrLf
                sql &= " FROM class_classinfo cc" & vbCrLf
                sql &= " JOIN plan_planinfo pp on pp.planid =cc.planid and pp.comidno=cc.comidno and pp.seqno =cc.seqno" & vbCrLf
                sql &= " JOIN org_orginfo oo on oo.comidno =cc.comidno" & vbCrLf
                sql &= " JOIN view_plan ip on ip.planid=cc.planid" & vbCrLf
                sql &= " JOIN Class_StudentsOfClass cs on cc.ocid =cs.ocid" & vbCrLf
                sql &= " LEFT JOIN WKRT2 ON WKRT2.RTReasonID=cs.RTReasonID" & vbCrLf
                sql &= " LEFT JOIN WKRT3 ON WKRT3.RTReasonID=cs.RTReasonID" & vbCrLf
                sql &= " LEFT JOIN WKRT6 ON WKRT6.RTReasonID=cs.RTReasonID" & vbCrLf

                sql &= " LEFT JOIN Stud_GetJobState3 sg3 on sg3.socid =cs.socid and sg3.CPoint=1 and sg3.IsGetJob=1" & vbCrLf
                sql &= " LEFT JOIN Stud_GetJobState3 sg9 on sg9.socid =cs.socid and sg9.CPoint= 9" & vbCrLf
                sql &= " where 1=1" & vbCrLf
                sql &= " and cc.NotOpen='N'" & vbCrLf
                sql &= " and cc.IsSuccess='Y'" & vbCrLf
                sql &= " and cc.FTDate < getdate()" & vbCrLf
                sql &= " and cs.MakeSOCID is null" & vbCrLf
                'sql &= " AND ip.years ='2016'" & vbCrLf
                'sql &= " AND ip.TPLANID ='02'" & vbCrLf
                If Identity1 <> "" Then
                    sql &= " and cs.MIdentityID IN (" & Identity1 & ")" & vbCrLf
                End If
                If sBudID <> "" Then
                    sql &= " and cs.BudgetID IN (" & sBudID & ")" & vbCrLf
                End If
                If PlanID.Value <> "" Then
                    sql &= "  and cc.PlanID ='" & PlanID.Value & "'" & vbCrLf
                End If
                If RIDValue.Value <> "" Then
                    sql &= "  and cc.RID ='" & RIDValue.Value & "'" & vbCrLf
                End If
                If OCIDStr <> "" Then
                    sql &= "  and cc.OCID IN (" & OCIDStr & ")" & vbCrLf
                End If
                If Syear.SelectedValue <> "" Then
                    sql &= "  and ip.years ='" & Syear.SelectedValue & "'" & vbCrLf
                End If
                If DistID1 <> "" Then
                    sql &= "  and ip.DistID IN (" & DistID1 & ")" & vbCrLf
                End If
                If TPlanID1 <> "" Then
                    sql &= "  and ip.TPlanID IN (" & TPlanID1 & ")" & vbCrLf
                End If
                If STDate1.Text <> "" Then
                    sql &= "  and cc.STDate>=" & TIMS.To_date(STDate1.Text) & vbCrLf
                End If
                If STDate2.Text <> "" Then
                    sql &= "  and cc.STDate<=" & TIMS.To_date(STDate2.Text) & vbCrLf
                End If
                If FTDate1.Text <> "" Then
                    sql &= "  and cc.FTDate>=" & TIMS.To_date(FTDate1.Text) & vbCrLf
                End If
                If FTDate2.Text <> "" Then
                    sql &= "  and cc.FTDate<=" & TIMS.To_date(FTDate2.Text) & vbCrLf
                End If
                sql &= " )" & vbCrLf
                sql &= " ,WGS2 AS (" & vbCrLf
                sql &= " select a.OCID" & vbCrLf
                sql &= " ,COUNT(a.opencount) opencount" & vbCrLf
                sql &= " ,COUNT(a.closecount) closecount" & vbCrLf
                sql &= " ,COUNT(a.jobcount) jobcount" & vbCrLf
                sql &= " ,COUNT(a.JobrelxCount) JobrelxCount" & vbCrLf
                sql &= " ,COUNT(a.x04) x04" & vbCrLf
                sql &= " ,COUNT(a.x03) x03" & vbCrLf
                sql &= " ,COUNT(a.x07) x07" & vbCrLf
                sql &= " ,COUNT(a.x31) x31" & vbCrLf
                sql &= " ,COUNT(a.x32) x32" & vbCrLf

                sql &= " ,COUNT(a.x02) x02" & vbCrLf
                sql &= " ,COUNT(a.x20) x20" & vbCrLf
                sql &= " ,COUNT(a.x21) x21" & vbCrLf
                sql &= " ,COUNT(a.x22) x22" & vbCrLf
                sql &= " ,COUNT(a.x98) x98" & vbCrLf
                sql &= " ,COUNT(a.j9x3) j9x3" & vbCrLf
                sql &= " ,COUNT(a.j9x2) j9x2" & vbCrLf
                sql &= " ,COUNT(a.j9x1) j9x1" & vbCrLf
                sql &= " " & vbCrLf
                sql &= " ,COUNT(a.x01) x01" & vbCrLf
                sql &= " ,COUNT(a.x13) x13" & vbCrLf
                sql &= " ,COUNT(a.x14) x14" & vbCrLf
                sql &= " ,COUNT(a.x99) x99" & vbCrLf
                sql &= " ,COUNT(a.xALL) xALL" & vbCrLf
                sql &= " ,COUNT(a.sum_PUBLICRESCUE) sum_PUBLICRESCUE" & vbCrLf
                sql &= " ,COUNT(a.sum_PUBLICRESCUE9) sum_PUBLICRESCUE9" & vbCrLf
                sql &= " ,COUNT(a.sum_jobrelx) sum_jobrelx" & vbCrLf
                sql &= " FROM WGS1 a" & vbCrLf
                sql &= " GROUP BY a.OCID" & vbCrLf
                sql &= " )" & vbCrLf
                'sql &= " SELECT WTD1.TDistName" & vbCrLf
                'sql &= " ,WTP1.TPlanName" & vbCrLf
                'sql &= " ,WID1.IDName" & vbCrLf
                sql &= " SELECT cc.years" & vbCrLf
                sql &= " ,a.OCID" & vbCrLf
                sql &= " ,cc.distid,cc.distname" & vbCrLf
                sql &= " ,cc.tplanid,cc.planname" & vbCrLf
                sql &= " ,cc.orgid,cc.orgname" & vbCrLf
                sql &= " ,cc.classcname2" & vbCrLf
                sql &= " ,CONVERT(varchar, cc.stdate, 111) stdate" & vbCrLf
                sql &= " ,CONVERT(varchar, cc.ftdate, 111) ftdate" & vbCrLf
                sql &= " ,a.opencount" & vbCrLf
                sql &= " ,a.closecount" & vbCrLf
                sql &= " ,a.jobcount" & vbCrLf
                sql &= " ,a.JobrelxCount" & vbCrLf
                sql &= " ,a.x04" & vbCrLf
                sql &= " ,a.x03" & vbCrLf
                sql &= " ,a.x07" & vbCrLf
                sql &= " ,a.x31" & vbCrLf
                sql &= " ,a.x32" & vbCrLf

                sql &= " ,a.x02" & vbCrLf
                sql &= " ,a.x20" & vbCrLf
                sql &= " ,a.x21" & vbCrLf
                sql &= " ,a.x22" & vbCrLf
                sql &= " ,a.x98" & vbCrLf
                sql &= " ,a.j9x3" & vbCrLf
                sql &= " ,a.j9x2" & vbCrLf
                sql &= " ,a.j9x1" & vbCrLf
                sql &= " ,a.x01" & vbCrLf
                sql &= " ,a.x13" & vbCrLf
                sql &= " ,a.x14" & vbCrLf
                sql &= " ,a.x99" & vbCrLf
                sql &= " ,a.xALL" & vbCrLf
                sql &= " ,a.sum_PUBLICRESCUE" & vbCrLf
                sql &= " ,a.sum_PUBLICRESCUE9" & vbCrLf
                sql &= " ,a.sum_jobrelx" & vbCrLf
                sql &= " FROM WGS2 a" & vbCrLf
                sql &= " JOIN VIEW2 cc on cc.ocid =a.ocid" & vbCrLf
                'sql &= " CROSS JOIN WTD1" & vbCrLf
                'sql &= " CROSS JOIN WTP1" & vbCrLf
                'sql &= " CROSS JOIN WID1" & vbCrLf
                sql &= " ORDER BY cc.years,cc.DISTID,cc.orgname,cc.classcname" & vbCrLf

        End Select

        Call TIMS.OpenDbConn(objconn)
        Dim oCmd As New SqlCommand(sql, objconn)
        Dim dt1 As New DataTable
        With oCmd
            .Parameters.Clear()
            dt1.Load(.ExecuteReader())
        End With

        'sda.SelectCommand.Parameters.Clear()
        'TIMS.Fill(sql, sda, dt)

        Const cst_dcad_年度 As String = "年度"
        'Const cst_dcad_轄區代碼 As String = "轄區代碼"
        Const cst_dcad_轄區 As String = "轄區"

        Const cst_dcad_訓練計畫 As String = "訓練計畫"
        Const cst_dcad_培訓單位 As String = "培訓單位"
        Const cst_dcad_班別 As String = "班別"
        Const cst_dcad_開訓日期 As String = "開訓日期"
        Const cst_dcad_結訓日期 As String = "結訓日期"

        Const cst_dcad_開訓人數 As String = "開訓人數"
        Const cst_dcad_結訓人數 As String = "結訓人數"
        Const cst_dcad_就業人數 As String = "就業人數"
        Const cst_dcad_提前就業勞保勾稽人數 As String = "提前就業勞保切結人數"
        Const cst_dcad_提前就業學員切結人數 As String = "提前就業學員切結人數"
        Const cst_dcad_提前就業雇主切結人數 As String = "提前就業雇主切結人數"
        Const cst_dcad_提前就業小計 As String = "提前就業小計"

        Const cst_dcad_就業關聯人數 As String = "就業關聯人數"
        Const cst_dcad_患病或遇意外傷害 As String = "患病或遇意外傷害"
        Const cst_dcad_遇家庭等災變事故 As String = "遇家庭等災變事故"
        Const cst_dcad_自願接受徵集入營者 As String = "自願、接受徵集入營者" 'X7
        Const cst_dcad_工作異動 As String = "工作異動" 'X31
        Const cst_dcad_課程內容不符預期 As String = "課程內容不符預期" 'X32

        'Const cst_dcad_其他離訓 As String = "其他(離訓)"
        Const cst_dcad_職類適性不合 As String = "職類適性不合" ' ,a.x20
        Const cst_dcad_經濟因素 As String = "經濟因素" ',a.x21
        Const cst_dcad_生涯規劃 As String = "生涯規劃" ',a.x22
        Const cst_dcad_其他經專案核准 As String = "其他經專案核准(出國、待產...)" ',a.x98

        Const cst_dcad_缺課時數超過規定 As String = "缺課時數超過規定"
        Const cst_dcad_參訓期間行為不檢情節重大 As String = "參訓期間行為不檢情節重大"
        Const cst_dcad_訓期未滿12找到工作 As String = "訓期未滿1/2找到工作"
        Const cst_dcad_其他退訓 As String = "其他(退訓)"
        Const cst_dcad_合計 As String = "合計"

        'https://msdn.microsoft.com/zh-tw/library/system.data.datacolumn.datatype(v=vs.110).aspx
        Dim dt As New DataTable
        Select Case sType
            Case 1
                dt.Columns.Add(New DataColumn(cst_dcad_年度))
                'dt.Columns.Add(New DataColumn(cst_dcad_轄區代碼))
                dt.Columns.Add(New DataColumn(cst_dcad_轄區))

            Case Else
                dt.Columns.Add(New DataColumn(cst_dcad_年度))
                'dt.Columns.Add(New DataColumn(cst_dcad_轄區代碼))
                dt.Columns.Add(New DataColumn(cst_dcad_轄區))
                dt.Columns.Add(New DataColumn(cst_dcad_訓練計畫))
                dt.Columns.Add(New DataColumn(cst_dcad_培訓單位))
                dt.Columns.Add(New DataColumn(cst_dcad_班別))
                dt.Columns.Add(New DataColumn(cst_dcad_開訓日期))
                dt.Columns.Add(New DataColumn(cst_dcad_結訓日期))

        End Select
        dt.Columns.Add(New DataColumn(cst_dcad_開訓人數, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_結訓人數, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_就業人數, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_提前就業勞保勾稽人數, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_提前就業學員切結人數, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_提前就業雇主切結人數, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_提前就業小計, System.Type.GetType("System.Int32")))

        dt.Columns.Add(New DataColumn(cst_dcad_就業關聯人數, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_患病或遇意外傷害, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_遇家庭等災變事故, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_自願接受徵集入營者, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_工作異動, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_課程內容不符預期, System.Type.GetType("System.Int32")))

        dt.Columns.Add(New DataColumn(cst_dcad_職類適性不合, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_經濟因素, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_生涯規劃, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_其他經專案核准, System.Type.GetType("System.Int32")))

        dt.Columns.Add(New DataColumn(cst_dcad_缺課時數超過規定, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_參訓期間行為不檢情節重大, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_訓期未滿12找到工作, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_其他退訓, System.Type.GetType("System.Int32")))
        dt.Columns.Add(New DataColumn(cst_dcad_合計, System.Type.GetType("System.Int32")))

        '合計
        If dt1.Rows.Count > 0 Then
            'Dim dr As DataRow
            '開訓人數	
            '結訓人數	
            '就業人數	
            '提前就業人數	
            '缺課時數超過規定	
            '遇家庭等災變事故	
            '患病或遇意外傷害	
            '適應困難	
            '訓練成績不合格	
            '奉召服兵役	
            '升學	
            '找到工作	
            '職類適性不合	
            '經濟因素	
            '進修	
            '其他(出國、待產....)

            Dim int_開訓人數 As Integer = 0
            Dim int_結訓人數 As Integer = 0
            Dim int_就業人數 As Integer = 0
            Dim int_提前就業_勞保勾稽人數 As Integer = 0
            Dim int_提前就業_學員切結人數 As Integer = 0
            Dim int_提前就業_雇主切結人數 As Integer = 0
            Dim int_提前就業_小計 As Integer = 0
            Dim int_就業關聯人數 As Integer = 0

            Dim int_X04 As Integer = 0
            Dim int_X03 As Integer = 0
            Dim int_X07 As Integer = 0
            Dim int_X31 As Integer = 0
            Dim int_X32 As Integer = 0

            'Dim int_X98 As Integer = 0
            Dim int_X01 As Integer = 0
            Dim int_X13 As Integer = 0
            Dim int_X14 As Integer = 0
            Dim int_X99 As Integer = 0
            Dim int_合計 As Integer = 0

            Dim dr As DataRow = Nothing
            For Each dr1 As DataRow In dt1.Rows
                int_開訓人數 += dr1("opencount")
                int_結訓人數 += dr1("closecount")
                int_就業人數 += dr1("jobcount")
                int_提前就業_勞保勾稽人數 += dr1("J9X3")
                int_提前就業_學員切結人數 += dr1("J9X2")
                int_提前就業_雇主切結人數 += dr1("J9X1")
                int_提前就業_小計 += (dr1("J9X3") + dr1("J9X2") + dr1("J9X1"))

                int_就業關聯人數 += dr1("JobrelxCount")
                int_X04 += dr1("X04")
                int_X03 += dr1("X03")
                int_X07 += dr1("X07")
                int_X31 += dr1("X31")
                int_X32 += dr1("X32")

                'int_X98 += dr1("X98")
                int_X01 += dr1("X01")
                int_X13 += dr1("X13")
                int_X14 += dr1("X14")
                int_X99 += dr1("X99")
                int_合計 += dr1("xALL")

                dr = dt.NewRow()
                dr(cst_dcad_年度) = dr1("Years")
                dr(cst_dcad_轄區) = dr1("distname")

                Select Case sType
                    Case 1
                        dr(cst_dcad_年度) = dr1("Years")
                        dr(cst_dcad_轄區) = dr1("distname")

                    Case Else
                        dr(cst_dcad_年度) = dr1("Years")
                        dr(cst_dcad_轄區) = dr1("distname")
                        dr(cst_dcad_訓練計畫) = dr1("planname")
                        dr(cst_dcad_培訓單位) = dr1("orgname")
                        dr(cst_dcad_班別) = dr1("classcname2")
                        dr(cst_dcad_開訓日期) = dr1("stdate")
                        dr(cst_dcad_結訓日期) = dr1("ftdate")

                End Select

                dr(cst_dcad_開訓人數) = dr1("opencount")
                dr(cst_dcad_結訓人數) = dr1("closecount")
                dr(cst_dcad_就業人數) = dr1("jobcount")
                dr(cst_dcad_提前就業勞保勾稽人數) = dr1("J9X3")
                dr(cst_dcad_提前就業學員切結人數) = dr1("J9X2")
                dr(cst_dcad_提前就業雇主切結人數) = dr1("J9X1")
                dr(cst_dcad_提前就業小計) = (dr1("J9X3") + dr1("J9X2") + dr1("J9X1"))

                dr(cst_dcad_就業關聯人數) = dr1("JobrelxCount")
                'dr("就業關聯率") = dub_就業關聯率
                dr(cst_dcad_患病或遇意外傷害) = dr1("X04")
                dr(cst_dcad_遇家庭等災變事故) = dr1("X03")
                dr(cst_dcad_自願接受徵集入營者) = dr1("X07")
                dr(cst_dcad_工作異動) = dr1("X31")
                dr(cst_dcad_課程內容不符預期) = dr1("X32")

                'dr("其他(離訓)") = dr1("X98")
                dr(cst_dcad_缺課時數超過規定) = dr1("X01")
                dr(cst_dcad_參訓期間行為不檢情節重大) = dr1("X13")
                dr(cst_dcad_訓期未滿12找到工作) = dr1("X14")
                dr(cst_dcad_其他退訓) = dr1("X99")
                dr(cst_dcad_合計) = dr1("xALL")

                dt.Rows.Add(dr)
            Next

            'dub_就業關聯率 = 0
            'If int_就業人數 > 0 Then
            '    dub_就業關聯率 = TIMS.Round(CDbl(int_就業關聯人數 / int_就業人數), 2)
            'End If

            '最後1行加總
            'sType 1:統計資料 2:明細資料
            dr = dt.NewRow()
            Select Case sType
                Case 1
                    dr(cst_dcad_年度) = cst_dcad_合計
                    dr(cst_dcad_轄區) = " "
                Case 2
                    dr(cst_dcad_年度) = cst_dcad_合計
                    dr(cst_dcad_轄區) = " "
                    dr(cst_dcad_訓練計畫) = " "
                    dr(cst_dcad_培訓單位) = " "
                    dr(cst_dcad_班別) = " "
                    dr(cst_dcad_開訓日期) = " "
                    dr(cst_dcad_結訓日期) = " "
            End Select

            dr(cst_dcad_開訓人數) = int_開訓人數
            dr(cst_dcad_結訓人數) = int_結訓人數
            dr(cst_dcad_就業人數) = int_就業人數
            dr(cst_dcad_提前就業勞保勾稽人數) = int_提前就業_勞保勾稽人數
            dr(cst_dcad_提前就業學員切結人數) = int_提前就業_學員切結人數
            dr(cst_dcad_提前就業雇主切結人數) = int_提前就業_雇主切結人數
            dr(cst_dcad_提前就業小計) = int_提前就業_小計

            dr(cst_dcad_就業關聯人數) = int_就業關聯人數
            'dr("就業關聯率") = dub_就業關聯率
            dr(cst_dcad_患病或遇意外傷害) = int_X04
            dr(cst_dcad_遇家庭等災變事故) = int_X03
            dr(cst_dcad_自願接受徵集入營者) = int_X07
            'dr("其他(離訓)") = dr1("X98")int_X98

            dr(cst_dcad_缺課時數超過規定) = int_X01
            dr(cst_dcad_參訓期間行為不檢情節重大) = int_X13
            dr(cst_dcad_訓期未滿12找到工作) = int_X14
            dr(cst_dcad_其他退訓) = int_X99
            dr(cst_dcad_合計) = int_合計

            dt.Rows.Add(dr)
        End If
        dt.AcceptChanges()

        'For i As Integer = 0 To dt.Columns.Count - 1
        '    Select Case dt.Columns(i).ColumnName
        '        Case "X04"
        '            dt.Columns(i).ColumnName = "患病或遇意外傷害"
        '        Case "X03"
        '            dt.Columns(i).ColumnName = "遇家庭等災變事故"
        '        Case "X07"
        '            dt.Columns(i).ColumnName = "自願、接受徵集入營者"
        '        Case "X98"
        '            dt.Columns(i).ColumnName = "其他(離訓)"
        '        Case "X01"
        '            dt.Columns(i).ColumnName = "缺課時數超過規定"
        '        Case "X13"
        '            dt.Columns(i).ColumnName = "參訓期間行為不檢情節重大"
        '        Case "X14"
        '            dt.Columns(i).ColumnName = "訓期未滿1/2找到工作"
        '        Case "X99"
        '            dt.Columns(i).ColumnName = "其他(退訓)"
        '    End Select
        '    'dt.Columns(i).ColumnName &= i.ToString()
        'Next

        ExportMsg.Text = "查無資料"
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            ExportMsg.Text = ""

            With DataGrid1
                .Visible = True
                .DataSource = dt
                .DataBind()
            End With

        End If
    End Sub


    '查詢 (匯出Excel用) [SQL] 2016
    Private Sub Search2(ByVal sType As Integer)
        'sType 1:統計資料 2:明細資料
        Dim OCIDStr As String = ""
        Dim TPlanID1 As String = ""
        Dim DistID1 As String = ""
        Dim Identity1 As String = ""
        Dim sBudID As String = ""

        '身分別參數
        Identity1 = ""
        For i As Integer = 1 To Me.Identity.Items.Count - 1
            If Me.Identity.Items(i).Selected Then
                If Identity1 <> "" Then Identity1 += ","
                Identity1 += Convert.ToString("'" & Me.Identity.Items(i).Value & "'")
            End If
        Next

        '預算別
        sBudID = ""
        For i As Integer = 0 To Me.BudID.Items.Count - 1
            If Me.BudID.Items(i).Selected Then
                If sBudID <> "" Then sBudID += ","
                sBudID += Convert.ToString("'" & Me.BudID.Items(i).Value & "'")
            End If
        Next

        '轄區參數
        DistID1 = ""
        'DistName = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 += ","
                DistID1 += Convert.ToString("'" & Me.DistID.Items(i).Value & "'")
            End If
        Next

        '訓練計畫參數
        TPlanID1 = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected Then
                If TPlanID1 <> "" Then TPlanID1 += ","
                TPlanID1 += Convert.ToString("'" & Me.TPlanID.Items(i).Value & "'")
            End If
        Next

        OCIDStr = ""
        For Each item As ListItem In OCID.Items
            If item.Selected = True Then
                If item.Value = "%" Then '有選擇全選
                    OCIDStr = ""
                    For i As Integer = 1 To Me.OCID.Items.Count - 1
                        If OCIDStr <> "" Then OCIDStr += ","
                        OCIDStr += Convert.ToString(Me.OCID.Items(i).Value)
                    Next
                    Exit For
                Else
                    If OCIDStr <> "" Then OCIDStr += ","
                    OCIDStr += item.Value
                End If
            End If
        Next

        '勾選班級後會省略結訓日期的條件
        If OCIDStr <> "" Then
            If (RIDValue.Value <> "") OrElse (PlanID.Value <> "") Then
                FTDate1.Text = ""
                FTDate2.Text = ""

                STDate1.Text = ""
                STDate2.Text = ""
                Syear.SelectedIndex = -1
            End If
        End If


        Dim sql As String = ""
        'sType 1:統計資料 2:明細資料
        Select Case sType
            Case 1
                sql = "" & vbCrLf
                sql += " WITH WKRT2 AS (SELECT SORT2,RTReasonID,Reason FROM Key_RejectTReason WHERE SORT2 IS NOT NULL ORDER BY SORT2)" & vbCrLf
                sql += " ,WKRT3 AS (SELECT SORT3,RTReasonID,Reason FROM Key_RejectTReason WHERE SORT3 IS NOT NULL ORDER BY SORT3)" & vbCrLf
                sql += " SELECT ip.years 年度" & vbCrLf
                sql += " ,ip.distid 轄區代碼" & vbCrLf
                sql += " ,ip.distname 轄區" & vbCrLf

                sql += " ,SUM(dbo.NVL(gcs.opencount,0)) 開訓人數" & vbCrLf
                sql += " ,SUM(dbo.NVL(gcs.closecount,0)) 結訓人數" & vbCrLf
                sql += " ,SUM(dbo.NVL(gcs.jobcount,0)) 就業人數" & vbCrLf
                sql += " ,SUM(dbo.NVL(gcs.x2,0)) 提前就業人數" & vbCrLf
                '就業關聯人數
                sql += " ,SUM(dbo.NVL(gcs.JobrelxCount,0)) 就業關聯人數" & vbCrLf
                ''就業關聯率
                'sql += " ,AVG(dbo.fn_GetJobRelRate(dbo.NVL(gcs.jobcount,0),dbo.NVL(gcs.JobrelxCount,0))) 就業關聯率" & vbCrLf
                'sql += " ,SUM(dbo.NVL(gcs.x4,0)) 患病或遇意外傷害" & vbCrLf
                'sql += " ,SUM(dbo.NVL(gcs.x3,0)) 遇家庭等災變事故" & vbCrLf
                'sql += " ,SUM(dbo.NVL(gcs.x7,0)) 奉召服兵役" & vbCrLf
                'sql += " ,SUM(dbo.NVL(gcs.x98,0)) 其他離訓" & vbCrLf
                'sql += " ,SUM(dbo.NVL(gcs.x1,0)) 缺課時數超過規定" & vbCrLf
                'sql += " ,SUM(dbo.NVL(gcs.x13,0)) 參訓期間行為不檢情節重大" & vbCrLf
                'sql += " ,SUM(dbo.NVL(gcs.x14,0)) 訓期未滿1/2找到工作" & vbCrLf
                'sql += " ,SUM(dbo.NVL(gcs.x99,0)) 其他退訓" & vbCrLf
                'sql += " ,SUM(dbo.NVL(gcs.xALL,0)) 合計" & vbCrLf
                sql += " ,SUM(dbo.NVL(gcs.x4,0)) X4" & vbCrLf
                sql += " ,SUM(dbo.NVL(gcs.x3,0)) X3" & vbCrLf
                sql += " ,SUM(dbo.NVL(gcs.x7,0)) X7" & vbCrLf
                sql += " ,SUM(dbo.NVL(gcs.x98,0)) X98" & vbCrLf
                sql += " ,SUM(dbo.NVL(gcs.x1,0)) X1" & vbCrLf
                sql += " ,SUM(dbo.NVL(gcs.x13,0)) X13" & vbCrLf
                sql += " ,SUM(dbo.NVL(gcs.x14,0)) X14" & vbCrLf
                sql += " ,SUM(dbo.NVL(gcs.x99,0)) X99" & vbCrLf
                sql += " ,SUM(dbo.NVL(gcs.xALL,0)) 合計" & vbCrLf
                'table.Columns(i).ColumnName &= i.ToString()
            Case 2
                sql = "" & vbCrLf
                sql += " WITH WKRT2 AS (SELECT SORT2,RTReasonID,Reason FROM Key_RejectTReason WHERE SORT2 IS NOT NULL ORDER BY SORT2)" & vbCrLf
                sql += " ,WKRT3 AS (SELECT SORT3,RTReasonID,Reason FROM Key_RejectTReason WHERE SORT3 IS NOT NULL ORDER BY SORT3)" & vbCrLf
                sql += " SELECT ip.years  年度" & vbCrLf
                sql += " ,ip.distid 轄區代碼" & vbCrLf
                sql += " ,ip.distname 轄區" & vbCrLf

                sql += " ,ip.planname 訓練計畫" & vbCrLf
                sql += " ,oo.orgname 培訓單位" & vbCrLf
                sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) 班別" & vbCrLf
                sql += " ,CONVERT(varchar, cc.STDATE, 111) 開訓日期" & vbCrLf
                sql += " ,CONVERT(varchar, cc.FTDate, 111) 結訓日期" & vbCrLf

                sql += " ,dbo.NVL(gcs.opencount,0) 開訓人數" & vbCrLf
                sql += " ,dbo.NVL(gcs.closecount,0) 結訓人數" & vbCrLf
                sql += " ,dbo.NVL(gcs.jobcount,0) 就業人數" & vbCrLf
                sql += " ,dbo.NVL(gcs.x2,0) 提前就業人數" & vbCrLf
                '就業關聯人數
                sql += " ,dbo.NVL(gcs.JobrelxCount,0) 就業關聯人數" & vbCrLf
                ''就業關聯率
                'sql += " ,dbo.fn_GetJobRelRate(dbo.NVL(gcs.jobcount,0),dbo.NVL(gcs.JobrelxCount,0)) 就業關聯率" & vbCrLf
                sql += " ,dbo.NVL(gcs.x4,0) X4" & vbCrLf
                sql += " ,dbo.NVL(gcs.x3,0) X3" & vbCrLf
                sql += " ,dbo.NVL(gcs.x7,0) X7" & vbCrLf
                sql += " ,dbo.NVL(gcs.x98,0) X98" & vbCrLf
                sql += " ,dbo.NVL(gcs.x1,0) X1" & vbCrLf
                sql += " ,dbo.NVL(gcs.x13,0) X13" & vbCrLf
                sql += " ,dbo.NVL(gcs.x14,0) X14" & vbCrLf
                sql += " ,dbo.NVL(gcs.x99,0) X99" & vbCrLf
                sql += " ,dbo.NVL(gcs.xALL,0) 合計" & vbCrLf
        End Select

        sql += " from class_classinfo cc" & vbCrLf
        sql += " join plan_planinfo pp on pp.planid =cc.planid and pp.comidno=cc.comidno and pp.seqno =cc.seqno" & vbCrLf
        sql += " and cc.NotOpen='N' and cc.IsSuccess='Y' " & vbCrLf
        sql += " join org_orginfo oo on oo.comidno =cc.comidno" & vbCrLf
        sql += " join view_plan ip on ip.planid=cc.planid " & vbCrLf
        sql += " left join (" & vbCrLf

        sql += "    select cs.ocid" & vbCrLf
        sql += "    ,count(1) opencount" & vbCrLf
        sql += "    ,count(case when cs.Studstatus not in (2,3) then 1 END) closecount" & vbCrLf
        sql += "    ,count(case when cs.Studstatus not in (2,3) and j3.socid is not null then 1 end ) jobcount" & vbCrLf
        sql += "    ,COUNT(case when cs.Studstatus not in (2,3) AND j3.JOBRELATE='Y' then 1 end ) JobrelxCount" & vbCrLf
        'SELECT * FROM KEY_REJECTTREASON WHERE ROWNUM <=10
        '提前就業人數
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='02' then 1 end) x2" & vbCrLf

        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='01' then 1 end) x1" & vbCrLf
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='03' then 1 end) x3" & vbCrLf
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='04' then 1 end) x4" & vbCrLf
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='05' then 1 end) x5" & vbCrLf
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='06' then 1 end) x6" & vbCrLf
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='07' then 1 end) x7" & vbCrLf
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='08' then 1 end) x8" & vbCrLf
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='09' then 1 end) x9" & vbCrLf
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='10' then 1 end) x10" & vbCrLf
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='11' then 1 end) x11" & vbCrLf
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='12' then 1 end) x12" & vbCrLf
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='13' then 1 end) x13" & vbCrLf
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='14' then 1 end) x14" & vbCrLf
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='98' then 1 end) x98" & vbCrLf
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and cs.RTReasonID ='99' then 1 end) x99" & vbCrLf
        sql += " 	,COUNT(case when cs.Studstatus in (2,3) and (WKRT2.RTReasonID is not null or WKRT3.RTReasonID is not null) and  cs.RTReasonID !='02'  then 1 end) xALL" & vbCrLf
        sql += " 	FROM class_classinfo cc " & vbCrLf
        sql += " 	JOIN id_plan ip on ip.planid=cc.planid" & vbCrLf
        sql += "    JOIN Class_StudentsOfClass cs on cc.ocid =cs.ocid" & vbCrLf
        sql += "    LEFT JOIN WKRT2 ON WKRT2.RTReasonID=cs.RTReasonID" & vbCrLf
        sql += "    LEFT JOIN WKRT3 ON WKRT3.RTReasonID=cs.RTReasonID" & vbCrLf
        'Sql += " --	 (select * from Key_RejectTReason where rownum <=10) 離退原因
        sql += "    LEFT JOIN STUD_GETJOBSTATE3 j3 on j3.socid =cs.socid and j3.CPoint=1 and j3.IsGetJob=1" & vbCrLf
        sql += " 	where 1=1 " & vbCrLf
        sql += " 	and cc.FTDate < getdate() " & vbCrLf
        sql += " 	and cs.MakeSOCID is null" & vbCrLf
        If Identity1 <> "" Then
            sql += " and cs.MIdentityID IN (" & Identity1 & ")" & vbCrLf
        End If
        If sBudID <> "" Then
            sql += " and cs.BudgetID IN (" & sBudID & ")" & vbCrLf
        End If

        If PlanID.Value <> "" Then
            sql += "  and cc.PlanID ='" & PlanID.Value & "'" & vbCrLf
        End If
        If RIDValue.Value <> "" Then
            sql += "  and cc.RID ='" & RIDValue.Value & "'" & vbCrLf
        End If
        If OCIDStr <> "" Then
            sql += "  and cc.OCID  IN (" & OCIDStr & ")" & vbCrLf
        End If
        sql += " 	group by cs.ocid " & vbCrLf
        sql += " ) gcs on gcs.ocid =cc.ocid " & vbCrLf
        sql += " where 1=1" & vbCrLf
        If Syear.SelectedValue <> "" Then
            sql += "  and ip.years ='" & Syear.SelectedValue & "'" & vbCrLf
        End If
        If PlanID.Value <> "" Then
            sql += "  and ip.PlanID ='" & PlanID.Value & "'" & vbCrLf
            sql += "  and cc.PlanID ='" & PlanID.Value & "'" & vbCrLf
        End If
        If RIDValue.Value <> "" Then
            sql += "  and cc.RID ='" & RIDValue.Value & "'" & vbCrLf
        End If
        If OCIDStr <> "" Then
            sql += "  and cc.OCID  IN (" & OCIDStr & ")" & vbCrLf
        End If
        If DistID1 <> "" Then
            sql += "  and ip.DistID IN (" & DistID1 & ")" & vbCrLf
        End If
        If TPlanID1 <> "" Then
            sql += "  and ip.TPlanID IN (" & TPlanID1 & ")" & vbCrLf
        End If
        If Me.STDate1.Text <> "" Then
            sql += "  and cc.STDate>=convert(datetime, @STDate1, 111) " & vbCrLf
            'sda.SelectCommand.Parameters.Add("STDate1", SqlDbType.VarChar).Value = Me.STDate1.Text
        End If
        If Me.STDate2.Text <> "" Then
            sql += "  and cc.STDate<=convert(datetime, @STDate2, 111) " & vbCrLf
            'sda.SelectCommand.Parameters.Add("STDate2", SqlDbType.VarChar).Value = Me.STDate2.Text
        End If
        If Me.FTDate1.Text <> "" Then
            sql += "  and cc.FTDate>=convert(datetime, @FTDate1, 111) " & vbCrLf
            'sda.SelectCommand.Parameters.Add("FTDate1", SqlDbType.VarChar).Value = Me.FTDate1.Text
        End If
        If Me.FTDate2.Text <> "" Then
            sql += "  and cc.FTDate<=convert(datetime, @FTDate2, 111) " & vbCrLf
            'sda.SelectCommand.Parameters.Add("FTDate2", SqlDbType.VarChar).Value = Me.FTDate2.Text
        End If
        'If Me.RIDValue.Value <> "" Then
        '    sql += "  and cc.RID=@RID" & vbCrLf
        '    sda.SelectCommand.Parameters.Add("@RID", SqlDbType.VarChar).Value = Me.RIDValue.Value
        'End If

        'sType 1:統計資料 2:明細資料
        Select Case sType
            Case 1
                sql += " GROUP BY ip.years,ip.distid,ip.distname " & vbCrLf
                sql += " ORDER BY ip.years,ip.distid,ip.distname " & vbCrLf
            Case 2
                sql += " ORDER BY ip.years,ip.distid,ip.distname,oo.orgname,cc.classcname " & vbCrLf
        End Select

        Call TIMS.OpenDbConn(objconn)
        Dim oCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With oCmd
            .Parameters.Clear()
            '.Parameters.Add("STDate1", SqlDbType.VarChar).Value = Me.STDate1.Text
            If Me.STDate1.Text <> "" Then
                'sql += "  and cc.STDate>=convert(datetime, @STDate1, 111) " & vbCrLf
                .Parameters.Add("STDate1", SqlDbType.VarChar).Value = Me.STDate1.Text
            End If
            If Me.STDate2.Text <> "" Then
                'sql += "  and cc.STDate<=convert(datetime, @STDate2, 111) " & vbCrLf
                .Parameters.Add("STDate2", SqlDbType.VarChar).Value = Me.STDate2.Text
            End If
            If Me.FTDate1.Text <> "" Then
                'sql += "  and cc.FTDate>=convert(datetime, @FTDate1, 111) " & vbCrLf
                .Parameters.Add("FTDate1", SqlDbType.VarChar).Value = Me.FTDate1.Text
            End If
            If Me.FTDate2.Text <> "" Then
                'sql += "  and cc.FTDate<=convert(datetime, @FTDate2, 111) " & vbCrLf
                .Parameters.Add("FTDate2", SqlDbType.VarChar).Value = Me.FTDate2.Text
            End If
            dt.Load(.ExecuteReader())
        End With

        'sda.SelectCommand.Parameters.Clear()
        'TIMS.Fill(sql, sda, dt)

        '合計
        If dt.Rows.Count > 0 Then
            'Dim dr As DataRow
            '開訓人數	
            '結訓人數	
            '就業人數	
            '提前就業人數	
            '缺課時數超過規定	
            '遇家庭等災變事故	
            '患病或遇意外傷害	
            '適應困難	
            '訓練成績不合格	
            '奉召服兵役	
            '升學	
            '找到工作	
            '職類適性不合	
            '經濟因素	
            '進修	
            '其他(出國、待產....)

            Dim int_開訓人數 As Integer = 0
            Dim int_結訓人數 As Integer = 0
            Dim int_就業人數 As Integer = 0
            Dim int_提前就業人數 As Integer = 0
            Dim int_就業關聯人數 As Integer = 0
            'Dim dub_就業關聯率 As Double = 0
            'Dim int_缺課時數超過規定 As Integer = 0
            'Dim int_遇家庭等災變事故 As Integer = 0
            'Dim int_患病或遇意外傷害 As Integer = 0
            'Dim int_適應困難 As Integer = 0
            'Dim int_訓練成績不合格 As Integer = 0
            'Dim int_奉召服兵役 As Integer = 0
            'Dim int_升學 As Integer = 0
            'Dim int_找到工作 As Integer = 0
            'Dim int_職類適性不合 As Integer = 0
            'Dim int_經濟因素 As Integer = 0
            'Dim int_進修 As Integer = 0
            'Dim int_其他 As Integer = 0
            'SELECT SORT2,RTReasonID,Reason FROM Key_RejectTReason WHERE SORT2 IS NOT NULL ORDER BY SORT2
            '04患病或遇意外傷害
            '03遇家庭等災變事故
            '07奉召服兵役
            '--02提前就業(訓期滿1/2以上)
            '98其他(職前訓練須經分署/縣市政府專案認定)
            'SELECT SORT3,RTReasonID,Reason FROM Key_RejectTReason WHERE SORT3 IS NOT NULL ORDER BY SORT3
            '01缺課時數超過規定
            '13參訓期間行為不檢情節重大
            '14訓期未滿1/2找到工作
            '99其他
            Dim int_X4 As Integer = 0
            Dim int_X3 As Integer = 0
            Dim int_X7 As Integer = 0
            Dim int_X98 As Integer = 0
            Dim int_X1 As Integer = 0
            Dim int_X13 As Integer = 0
            Dim int_X14 As Integer = 0
            Dim int_X99 As Integer = 0
            Dim int_合計 As Integer = 0

            Dim dr As DataRow = Nothing
            For Each dr In dt.Rows
                int_開訓人數 += dr("開訓人數")
                int_結訓人數 += dr("結訓人數")
                int_就業人數 += dr("就業人數")
                int_提前就業人數 += dr("提前就業人數")
                int_就業關聯人數 += dr("就業關聯人數")
                int_X4 += dr("X4")
                int_X3 += dr("X3")
                int_X7 += dr("X7")
                int_X98 += dr("X98")
                int_X1 += dr("X1")
                int_X13 += dr("X13")
                int_X14 += dr("X14")
                int_X99 += dr("X99")
                int_合計 += dr("合計")
            Next

            'dub_就業關聯率 = 0
            'If int_就業人數 > 0 Then
            '    dub_就業關聯率 = TIMS.Round(CDbl(int_就業關聯人數 / int_就業人數), 2)
            'End If

            dr = dt.NewRow()

            '最後1行加總
            'sType 1:統計資料 2:明細資料
            Select Case sType
                Case 1
                    dr("轄區代碼") = " "
                    dr("轄區") = " "
                Case 2
                    dr("轄區代碼") = " "
                    dr("轄區") = " "
                    dr("訓練計畫") = " "
                    dr("培訓單位") = " "
                    dr("班別") = " "
                    dr("開訓日期") = " "
                    dr("結訓日期") = " "
            End Select

            dr("年度") = "合計"
            dr("開訓人數") = int_開訓人數
            dr("結訓人數") = int_結訓人數
            dr("就業人數") = int_就業人數
            dr("提前就業人數") = int_提前就業人數
            dr("就業關聯人數") = int_就業關聯人數
            'dr("就業關聯率") = dub_就業關聯率
            dr("X4") = int_X4
            dr("X3") = int_X3
            dr("X7") = int_X7
            dr("X98") = int_X98
            dr("X1") = int_X1
            dr("X13") = int_X13
            dr("X14") = int_X14
            dr("X99") = int_X99
            dr("合計") = int_合計

            dt.Rows.Add(dr)
        End If
        dt.AcceptChanges()

        For i As Integer = 0 To dt.Columns.Count - 1
            Select Case dt.Columns(i).ColumnName
                Case "X4"
                    dt.Columns(i).ColumnName = "患病或遇意外傷害"
                Case "X3"
                    dt.Columns(i).ColumnName = "遇家庭等災變事故"
                Case "X7"
                    dt.Columns(i).ColumnName = "奉召服兵役"
                Case "X98"
                    dt.Columns(i).ColumnName = "其他(離訓)"
                Case "X1"
                    dt.Columns(i).ColumnName = "缺課時數超過規定"
                Case "X13"
                    dt.Columns(i).ColumnName = "參訓期間行為不檢情節重大"
                Case "X14"
                    dt.Columns(i).ColumnName = "訓期未滿1/2找到工作"
                Case "X99"
                    dt.Columns(i).ColumnName = "其他(退訓)"
            End Select
            'dt.Columns(i).ColumnName &= i.ToString()
        Next

        ExportMsg.Text = "查無資料"
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            ExportMsg.Text = ""

            With DataGrid1
                .Visible = True
                .DataSource = dt
                .DataBind()
            End With

        End If
    End Sub

    '查詢 (匯出Excel用) [SQL] OLD
    Private Sub Search1(ByVal sType As Integer)
        'sType 1:統計資料 2:明細資料
        Dim OCIDStr As String = ""
        Dim TPlanID1 As String = ""
        Dim DistID1 As String = ""
        Dim Identity1 As String = ""
        Dim sBudID As String = ""

        '身分別參數
        Identity1 = ""
        For i As Integer = 1 To Me.Identity.Items.Count - 1
            If Me.Identity.Items(i).Selected Then
                If Identity1 <> "" Then Identity1 += ","
                Identity1 += Convert.ToString("'" & Me.Identity.Items(i).Value & "'")
            End If
        Next

        '預算別
        sBudID = ""
        For i As Integer = 0 To Me.BudID.Items.Count - 1
            If Me.BudID.Items(i).Selected Then
                If sBudID <> "" Then sBudID += ","
                sBudID += Convert.ToString("'" & Me.BudID.Items(i).Value & "'")
            End If
        Next

        '轄區參數
        DistID1 = ""
        'DistName = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 += ","
                DistID1 += Convert.ToString("'" & Me.DistID.Items(i).Value & "'")
            End If
        Next

        '訓練計畫參數
        TPlanID1 = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected Then
                If TPlanID1 <> "" Then TPlanID1 += ","
                TPlanID1 += Convert.ToString("'" & Me.TPlanID.Items(i).Value & "'")
            End If
        Next


        OCIDStr = ""
        For Each item As ListItem In OCID.Items
            If item.Selected = True Then
                If item.Value = "%" Then '有選擇全選
                    OCIDStr = ""
                    For i As Integer = 1 To Me.OCID.Items.Count - 1
                        If OCIDStr <> "" Then OCIDStr += ","
                        OCIDStr += Convert.ToString(Me.OCID.Items(i).Value)
                    Next
                    Exit For
                Else
                    If OCIDStr <> "" Then OCIDStr += ","
                    OCIDStr += item.Value
                End If
            End If
        Next

        '勾選班級後會省略結訓日期的條件
        If OCIDStr <> "" Then
            If (RIDValue.Value <> "") OrElse (PlanID.Value <> "") Then
                FTDate1.Text = ""
                FTDate2.Text = ""

                STDate1.Text = ""
                STDate2.Text = ""
                Syear.SelectedIndex = -1
            End If
        End If

        'Dim dt As DataTable = Nothing
        'Dim sda As SqlDataAdapter = Nothing
        'Try
        'Catch ex As Exception
        '    Dim vMsg As String = "系統錯誤：" & ex.Message.ToString
        '    ExportMsg.Text = vMsg
        '    'Common.MessageBox(Me, vMsg)
        'End Try
        'sda = TIMS.GetOneDA(objconn)

        Dim sql As String = ""
        'sType 1:統計資料 2:明細資料
        Select Case sType
            Case 1
                sql = "" & vbCrLf
                sql &= " SELECT ip.years  年度" & vbCrLf
                sql &= " ,ip.distid 轄區代碼" & vbCrLf
                sql &= " ,ip.distname 轄區" & vbCrLf
                sql &= " ,SUM(dbo.NVL(gcs.opencount,0)) 開訓人數" & vbCrLf
                sql &= " ,SUM(dbo.NVL(gcs.closecount,0)) 結訓人數" & vbCrLf
                sql &= " ,SUM(dbo.NVL(gcs.jobcount,0)) 就業人數" & vbCrLf
                sql &= " ,SUM(dbo.NVL(gcs.x2,0)) 提前就業人數" & vbCrLf
                '就業關聯人數
                sql &= " ,SUM(dbo.NVL(gcs.JobrelxCount,0)) 就業關聯人數" & vbCrLf
                ''就業關聯率
                'sql += " ,AVG(dbo.fn_GetJobRelRate(dbo.NVL(gcs.jobcount,0),dbo.NVL(gcs.JobrelxCount,0))) 就業關聯率" & vbCrLf
                sql &= " , SUM(dbo.NVL(gcs.x1,0)) 缺課時數超過規定" & vbCrLf
                sql &= " , SUM(dbo.NVL(gcs.x3,0)) 遇家庭等災變事故" & vbCrLf
                sql &= " , SUM(dbo.NVL(gcs.x4,0)) 患病或遇意外傷害" & vbCrLf
                sql &= " , SUM(dbo.NVL(gcs.x5,0)) 適應困難" & vbCrLf
                sql &= " , SUM(dbo.NVL(gcs.x6,0)) 訓練成績不合格" & vbCrLf
                sql &= " , SUM(dbo.NVL(gcs.x7,0)) 奉召服兵役" & vbCrLf
                sql &= " , SUM(dbo.NVL(gcs.x8,0)) 升學" & vbCrLf
                sql &= " , SUM(dbo.NVL(gcs.x9,0)) 找到工作" & vbCrLf
                sql &= " , SUM(dbo.NVL(gcs.x10,0)) 職類適性不合" & vbCrLf
                sql &= " , SUM(dbo.NVL(gcs.x11,0)) 經濟因素" & vbCrLf
                sql &= " , SUM(dbo.NVL(gcs.x12,0)) 進修" & vbCrLf
                sql &= " , SUM(dbo.NVL(gcs.x99,0)) 其他" & vbCrLf
            Case 2
                sql = "" & vbCrLf
                sql &= " SELECT ip.years  年度" & vbCrLf
                sql &= " ,ip.distid 轄區代碼" & vbCrLf
                sql &= " ,ip.distname 轄區" & vbCrLf
                sql &= " ,ip.planname 訓練計畫" & vbCrLf
                sql &= " ,oo.orgname 培訓單位" & vbCrLf
                sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) 班別" & vbCrLf
                sql &= " ,CONVERT(varchar, cc.STDATE, 111) 開訓日期" & vbCrLf
                sql &= " ,CONVERT(varchar, cc.FTDate, 111) 結訓日期" & vbCrLf
                sql &= " ,dbo.NVL(gcs.opencount,0) 開訓人數" & vbCrLf
                sql &= " ,dbo.NVL(gcs.closecount,0) 結訓人數" & vbCrLf
                sql &= " ,dbo.NVL(gcs.jobcount,0) 就業人數" & vbCrLf
                sql &= " ,dbo.NVL(gcs.x2,0) 提前就業人數" & vbCrLf
                '就業關聯人數
                sql &= " ,dbo.NVL(gcs.JobrelxCount,0) 就業關聯人數" & vbCrLf
                ''就業關聯率
                'sql += " ,dbo.fn_GetJobRelRate(dbo.NVL(gcs.jobcount,0),dbo.NVL(gcs.JobrelxCount,0)) 就業關聯率" & vbCrLf
                sql &= " ,dbo.NVL(gcs.x1,0) 缺課時數超過規定" & vbCrLf
                sql &= " ,dbo.NVL(gcs.x3,0) 遇家庭等災變事故" & vbCrLf
                sql &= " ,dbo.NVL(gcs.x4,0) 患病或遇意外傷害" & vbCrLf
                sql &= " ,dbo.NVL(gcs.x5,0) 適應困難" & vbCrLf
                sql &= " ,dbo.NVL(gcs.x6,0) 訓練成績不合格" & vbCrLf
                sql &= " ,dbo.NVL(gcs.x7,0) 奉召服兵役" & vbCrLf
                sql &= " ,dbo.NVL(gcs.x8,0) 升學" & vbCrLf
                sql &= " ,dbo.NVL(gcs.x9,0) 找到工作" & vbCrLf
                sql &= " ,dbo.NVL(gcs.x10,0) 職類適性不合" & vbCrLf
                sql &= " ,dbo.NVL(gcs.x11,0) 經濟因素" & vbCrLf
                sql &= " ,dbo.NVL(gcs.x12,0) 進修" & vbCrLf
                sql &= " ,dbo.NVL(gcs.x99,0) 其他" & vbCrLf
        End Select

        sql &= " from class_classinfo cc" & vbCrLf
        sql &= " join plan_planinfo pp on pp.planid =cc.planid and pp.comidno=cc.comidno and pp.seqno =cc.seqno" & vbCrLf
        sql &= " and cc.NotOpen='N' and cc.IsSuccess='Y' " & vbCrLf
        sql &= " join org_orginfo oo on oo.comidno =cc.comidno" & vbCrLf
        sql &= " join view_plan ip on ip.planid=cc.planid " & vbCrLf
        sql &= " left join (" & vbCrLf

        sql &= "    select cs.ocid" & vbCrLf
        sql &= "    ,count(1) opencount" & vbCrLf
        sql &= "    ,count(case when cs.Studstatus not in (2,3) then 1 END) closecount" & vbCrLf
        sql &= "    ,count(case when cs.Studstatus not in (2,3) and j3.socid is not null then 1 end ) jobcount" & vbCrLf
        sql &= "    ,COUNT(case when cs.Studstatus not in (2,3) AND j3.JOBRELATE='Y' then 1 end ) JobrelxCount" & vbCrLf
        'SELECT * FROM KEY_REJECTTREASON WHERE ROWNUM <=10
        '提前就業人數
        sql &= " 	,sum(case when cs.Studstatus in (2,3) and cs.RTReasonID ='02' then 1 end) x2" & vbCrLf

        sql &= " 	,sum(case when cs.Studstatus in (2,3) and cs.RTReasonID ='01' then 1 end) x1" & vbCrLf
        sql &= " 	,sum(case when cs.Studstatus in (2,3) and cs.RTReasonID ='03' then 1 end) x3" & vbCrLf
        sql &= " 	,sum(case when cs.Studstatus in (2,3) and cs.RTReasonID ='04' then 1 end) x4" & vbCrLf
        sql &= " 	,sum(case when cs.Studstatus in (2,3) and cs.RTReasonID ='05' then 1 end) x5" & vbCrLf
        sql &= " 	,sum(case when cs.Studstatus in (2,3) and cs.RTReasonID ='06' then 1 end) x6" & vbCrLf
        sql &= " 	,sum(case when cs.Studstatus in (2,3) and cs.RTReasonID ='07' then 1 end) x7" & vbCrLf
        sql &= " 	,sum(case when cs.Studstatus in (2,3) and cs.RTReasonID ='08' then 1 end) x8" & vbCrLf
        sql &= " 	,sum(case when cs.Studstatus in (2,3) and cs.RTReasonID ='09' then 1 end) x9" & vbCrLf
        sql &= " 	,sum(case when cs.Studstatus in (2,3) and cs.RTReasonID ='10' then 1 end) x10" & vbCrLf
        sql &= " 	,sum(case when cs.Studstatus in (2,3) and cs.RTReasonID ='11' then 1 end) x11" & vbCrLf
        sql &= " 	,sum(case when cs.Studstatus in (2,3) and cs.RTReasonID ='12' then 1 end) x12" & vbCrLf
        sql &= " 	,sum(case when cs.Studstatus in (2,3) and cs.RTReasonID ='99' then 1 end) x99" & vbCrLf
        sql &= " 	FROM class_classinfo cc " & vbCrLf
        sql &= " 	JOIN id_plan ip on ip.planid=cc.planid" & vbCrLf
        sql &= "    JOIN Class_StudentsOfClass cs on cc.ocid =cs.ocid" & vbCrLf
        'Sql += " --	 (select * from Key_RejectTReason where rownum <=10) 離退原因
        sql &= "    left join Stud_GetJobState3 j3 on j3.socid =cs.socid and j3.CPoint=1 and j3.IsGetJob=1" & vbCrLf
        sql &= " 	where 1=1 " & vbCrLf
        sql &= " 	and cc.FTDate < getdate() " & vbCrLf
        sql &= " 	and cs.MakeSOCID is null" & vbCrLf
        If Identity1 <> "" Then
            sql &= " and cs.MIdentityID IN (" & Identity1 & ")" & vbCrLf
        End If
        If sBudID <> "" Then
            sql &= " and cs.BudgetID IN (" & sBudID & ")" & vbCrLf
        End If

        If PlanID.Value <> "" Then
            sql &= "  and cc.PlanID ='" & PlanID.Value & "'" & vbCrLf
        End If
        If RIDValue.Value <> "" Then
            sql &= "  and cc.RID ='" & RIDValue.Value & "'" & vbCrLf
        End If
        If OCIDStr <> "" Then
            sql &= "  and cc.OCID  IN (" & OCIDStr & ")" & vbCrLf
        End If
        sql &= " 	group by cs.ocid " & vbCrLf
        sql &= " ) gcs on gcs.ocid =cc.ocid " & vbCrLf
        sql &= " where 1=1" & vbCrLf
        If Syear.SelectedValue <> "" Then
            sql &= "  and ip.years ='" & Syear.SelectedValue & "'" & vbCrLf
        End If
        If PlanID.Value <> "" Then
            sql &= "  and ip.PlanID ='" & PlanID.Value & "'" & vbCrLf
            sql &= "  and cc.PlanID ='" & PlanID.Value & "'" & vbCrLf
        End If
        If RIDValue.Value <> "" Then
            sql &= "  and cc.RID ='" & RIDValue.Value & "'" & vbCrLf
        End If
        If OCIDStr <> "" Then
            sql &= "  and cc.OCID  IN (" & OCIDStr & ")" & vbCrLf
        End If
        If DistID1 <> "" Then
            sql &= "  and ip.DistID IN (" & DistID1 & ")" & vbCrLf
        End If
        If TPlanID1 <> "" Then
            sql &= "  and ip.TPlanID IN (" & TPlanID1 & ")" & vbCrLf
        End If
        If Me.STDate1.Text <> "" Then
            sql &= "  and cc.STDate>=convert(datetime, @STDate1, 111) " & vbCrLf
            'sda.SelectCommand.Parameters.Add("STDate1", SqlDbType.VarChar).Value = Me.STDate1.Text
        End If
        If Me.STDate2.Text <> "" Then
            sql &= "  and cc.STDate<=convert(datetime, @STDate2, 111) " & vbCrLf
            'sda.SelectCommand.Parameters.Add("STDate2", SqlDbType.VarChar).Value = Me.STDate2.Text
        End If
        If Me.FTDate1.Text <> "" Then
            sql &= "  and cc.FTDate>=convert(datetime, @FTDate1, 111) " & vbCrLf
            'sda.SelectCommand.Parameters.Add("FTDate1", SqlDbType.VarChar).Value = Me.FTDate1.Text
        End If
        If Me.FTDate2.Text <> "" Then
            sql &= "  and cc.FTDate<=convert(datetime, @FTDate2, 111) " & vbCrLf
            'sda.SelectCommand.Parameters.Add("FTDate2", SqlDbType.VarChar).Value = Me.FTDate2.Text
        End If
        'If Me.RIDValue.Value <> "" Then
        '    sql &= "  and cc.RID=@RID" & vbCrLf
        '    sda.SelectCommand.Parameters.Add("@RID", SqlDbType.VarChar).Value = Me.RIDValue.Value
        'End If

        'sType 1:統計資料 2:明細資料
        Select Case sType
            Case 1
                sql &= " GROUP BY ip.years,ip.distid,ip.distname " & vbCrLf
                sql &= " ORDER BY ip.years,ip.distid,ip.distname " & vbCrLf
            Case 2
                sql &= " ORDER BY ip.years,ip.distid,ip.distname,oo.orgname,cc.classcname " & vbCrLf
        End Select

        Call TIMS.OpenDbConn(objconn)
        Dim oCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With oCmd
            .Parameters.Clear()
            '.Parameters.Add("STDate1", SqlDbType.VarChar).Value = Me.STDate1.Text
            If Me.STDate1.Text <> "" Then
                'sql += "  and cc.STDate>=convert(datetime, @STDate1, 111) " & vbCrLf
                .Parameters.Add("STDate1", SqlDbType.VarChar).Value = Me.STDate1.Text
            End If
            If Me.STDate2.Text <> "" Then
                'sql += "  and cc.STDate<=convert(datetime, @STDate2, 111) " & vbCrLf
                .Parameters.Add("STDate2", SqlDbType.VarChar).Value = Me.STDate2.Text
            End If
            If Me.FTDate1.Text <> "" Then
                'sql += "  and cc.FTDate>=convert(datetime, @FTDate1, 111) " & vbCrLf
                .Parameters.Add("FTDate1", SqlDbType.VarChar).Value = Me.FTDate1.Text
            End If
            If Me.FTDate2.Text <> "" Then
                'sql += "  and cc.FTDate<=convert(datetime, @FTDate2, 111) " & vbCrLf
                .Parameters.Add("FTDate2", SqlDbType.VarChar).Value = Me.FTDate2.Text
            End If
            dt.Load(.ExecuteReader())
        End With

        'sda.SelectCommand.Parameters.Clear()
        'TIMS.Fill(sql, sda, dt)

        '合計
        If dt.Rows.Count > 0 Then
            'Dim dr As DataRow
            '開訓人數	
            '結訓人數	
            '就業人數	
            '提前就業人數	
            '缺課時數超過規定	
            '遇家庭等災變事故	
            '患病或遇意外傷害	
            '適應困難	
            '訓練成績不合格	
            '奉召服兵役	
            '升學	
            '找到工作	
            '職類適性不合	
            '經濟因素	
            '進修	
            '其他(出國、待產....)

            Dim int_開訓人數 As Integer = 0
            Dim int_結訓人數 As Integer = 0
            Dim int_就業人數 As Integer = 0
            Dim int_提前就業人數 As Integer = 0
            Dim int_就業關聯人數 As Integer = 0
            'Dim dub_就業關聯率 As Double = 0
            Dim int_缺課時數超過規定 As Integer = 0
            Dim int_遇家庭等災變事故 As Integer = 0
            Dim int_患病或遇意外傷害 As Integer = 0
            Dim int_適應困難 As Integer = 0
            Dim int_訓練成績不合格 As Integer = 0
            Dim int_奉召服兵役 As Integer = 0
            Dim int_升學 As Integer = 0
            Dim int_找到工作 As Integer = 0
            Dim int_職類適性不合 As Integer = 0
            Dim int_經濟因素 As Integer = 0
            Dim int_進修 As Integer = 0
            Dim int_其他 As Integer = 0

            Dim dr As DataRow = Nothing
            For Each dr In dt.Rows
                int_開訓人數 += dr("開訓人數")
                int_結訓人數 += dr("結訓人數")
                int_就業人數 += dr("就業人數")
                int_提前就業人數 += dr("提前就業人數")
                int_就業關聯人數 += dr("就業關聯人數")
                int_缺課時數超過規定 += dr("缺課時數超過規定")
                int_遇家庭等災變事故 += dr("遇家庭等災變事故")
                int_患病或遇意外傷害 += dr("患病或遇意外傷害")
                int_適應困難 += dr("適應困難")
                int_訓練成績不合格 += dr("訓練成績不合格")
                int_奉召服兵役 += dr("奉召服兵役")
                int_升學 += dr("升學")
                int_找到工作 += dr("找到工作")
                int_職類適性不合 += dr("職類適性不合")
                int_經濟因素 += dr("經濟因素")
                int_進修 += dr("進修")
                int_其他 += dr("其他")
            Next

            'dub_就業關聯率 = 0
            'If int_就業人數 > 0 Then
            '    dub_就業關聯率 = TIMS.Round(CDbl(int_就業關聯人數 / int_就業人數), 2)
            'End If

            dr = dt.NewRow()

            'sType 1:統計資料 2:明細資料
            Select Case sType
                Case 1
                    dr("轄區代碼") = " "
                    dr("轄區") = " "
                Case 2
                    dr("轄區代碼") = " "
                    dr("轄區") = " "
                    dr("訓練計畫") = " "
                    dr("培訓單位") = " "
                    dr("班別") = " "
                    dr("開訓日期") = " "
                    dr("結訓日期") = " "
            End Select

            dr("年度") = "合計"
            dr("開訓人數") = int_開訓人數
            dr("結訓人數") = int_結訓人數
            dr("就業人數") = int_就業人數
            dr("提前就業人數") = int_提前就業人數
            dr("就業關聯人數") = int_就業關聯人數
            'dr("就業關聯率") = dub_就業關聯率
            dr("缺課時數超過規定") = int_缺課時數超過規定
            dr("遇家庭等災變事故") = int_遇家庭等災變事故
            dr("患病或遇意外傷害") = int_患病或遇意外傷害
            dr("適應困難") = int_適應困難
            dr("訓練成績不合格") = int_訓練成績不合格
            dr("奉召服兵役") = int_奉召服兵役
            dr("升學") = int_升學
            dr("找到工作") = int_找到工作
            dr("職類適性不合") = int_職類適性不合
            dr("經濟因素") = int_經濟因素
            dr("進修") = int_進修
            dr("其他") = int_其他
            dt.Rows.Add(dr)
        End If
        dt.AcceptChanges()

        ExportMsg.Text = "查無資料"
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            ExportMsg.Text = ""

            With DataGrid1
                .Visible = True
                .DataSource = dt
                .DataBind()
            End With

        End If

        'conn.Close()
        'If Not sda Is Nothing Then sda.Dispose()
        'If Not dt Is Nothing Then dt.Dispose()

    End Sub

#End Region

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯出Excel (dg匯出)
    Private Sub btnExport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport1.Click
        DataGrid1.AllowPaging = False '關閉分頁功能
        DataGrid1.EnableViewState = False  '把ViewState給關了

        'Dim flagtest As Boolean = TIMS.sUtl_ChkTest()
        Dim ssflagYears As String = ""
        Select Case searcha_type1.SelectedValue
            Case "1"
                '統計資料'查詢
                ssflagYears = cst_else_1
                If sm.UserInfo.Years >= "2016" Then ssflagYears = cst_2016_1 '2016
                If sm.UserInfo.Years >= "2017" Then ssflagYears = cst_2017_1 '2017
                If sm.UserInfo.Years >= "2020" Then ssflagYears = cst_2020_1 '2020

            Case "2"
                '明細資料'查詢
                ssflagYears = cst_else_2
                If sm.UserInfo.Years >= "2016" Then ssflagYears = cst_2016_2 '2016
                If sm.UserInfo.Years >= "2017" Then ssflagYears = cst_2017_2 '2017
                If sm.UserInfo.Years >= "2020" Then ssflagYears = cst_2020_2 '2020

        End Select

        Dim sFileName1 As String = "離退訓人數統計表"
        Select Case ssflagYears
            Case cst_2020_1
                Call Search4(1) '2020
            Case cst_2017_1
                '統計資料'查詢
                Call Search3(1) '2017
            Case cst_2016_1
                '統計資料'查詢
                Call Search2(1) '2016
            Case cst_else_1
                '統計資料'查詢
                Call Search1(1) 'old

            Case cst_2020_2
                sFileName1 = "離退訓人數統計"
                Call Search4(2) '2020
            Case cst_2017_2
                Call Search3(2) '2017
            Case cst_2016_2
                Call Search2(2) '2016
            Case cst_else_2
                Call Search1(2) 'old

            Case Else
                Common.MessageBox(Me, "查詢方式未選擇，請確認！")
                Exit Sub
        End Select

        If ExportMsg.Text <> "" Then
            Common.MessageBox(Page, ExportMsg.Text)
            Exit Sub
        End If

        'Dim strSTYLE As String = ""
        'strSTYLE &= ("<style>")
        DataGrid1.AllowPaging = False '關閉分頁功能
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)
        Dim strHTML As String = ""
        strHTML &= (TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        'parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'DataGrid1.Visible = False
        'Response.End()
    End Sub

End Class

