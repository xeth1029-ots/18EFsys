Partial Class SD_14_010
    Inherits AuthBasePage

    'ReportQuery,'SQControl.aspx,'/**OLD**/,'SD_14_010_07emp'空白變更申請書,'SD_14_010_07'列印,'/**old2**/,'SD_14_010_1'空白變更申請書(不要改),'SD_14_010_R1'列印變更課程表(不要改),'SD_14_010'列印(不要改),
    '/**NEW 2014**/,'SD_14_010*b.jrxml,'SD_14_010_1_b,'空白變更申請書,'SD_14_010_R1_b,'列印變更課程表,'SD_14_010_b,'列印資料內容
    'Const cst_printFN1 As String="SD_14_010_R1_b"  '列印變更課程表
    Public Const cst_printFN1c As String = "SD_14_010_R1_c"  '列印變更課程表-課程進度/內容字多版本
    Public Const cst_printFN1d As String = "SD_14_010_R1_d"  '列印變更課程表-課程進度/內容字多版本(增加-技檢訓練時數)
    Public Const cst_printFN2 As String = "SD_14_010_b"     '列印資料內容
    Public Const cst_printFN3 As String = "SD_14_010_1_b"   '空白變更申請書

    '技檢訓練時數 '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-01才存欄位，否清空。
    'Const cst_EHour_t1 As String="技檢訓練時數,目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時可儲存，若不符合上述條件，該資料不會存入資料庫。"
    '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-01才存欄位，否清空。
    'Const cst_EHour_Use_TMID As String="672"

    Const cst_ReviseStatus_txt_Y As String = "審核通過" '審核通過 AppliedResult Y
    'Const cst_ReviseStatus_txt_O As String = "審核後修正" '審核後修正 AppliedResult O
    Const cst_ReviseStatus_txt_N As String = "審核失敗" '審核失敗/審核不通過 AppliedResult N
    Const cst_ReviseStatus_txt_oth As String = "審核中"
    Const cst_ReviseStatus_txt_PARTREDUC_Y As String = "還原修正中" '"待修正" '待修正 AppliedResult oth
    Dim gsCmd As SqlCommand

    'Dim iPYNum14 As Integer=1 'TIMS.sUtl_GetPYNum14(Me)
    Dim prtFilename As String = "" '列印表件名稱

    Dim flag_PackageType_NOUSE As Boolean = True 'true:(未使用)未選擇 包班種類(移除) 'false:(使用)選擇 包班種類(保留)
    '產投使用／遠距教學 暫不啟用
    Dim flag_StopDISTANCE As Boolean = True

    Const Cst_變更項目 As Integer = 3
    Const Cst_審核狀態 As Integer = 4
    Dim ChgItemName As String()
    'Dim gda As SqlDataAdapter '= TIMS.GetOneDA()
    '產投選項，職前選項自行增減
    'ChgItem=TIMS.TPlanID28ChgItemName
    'ChgItem [直接改介面]
    Const Cst_i訓練期間 As Integer = 1
    Const Cst_i訓練時段 As Integer = 2
    Const Cst_i訓練地點 As Integer = 3
    Const Cst_i課程編配 As Integer = 4
    Const Cst_i訓練師資 As Integer = 5
    Const Cst_i班別名稱 As Integer = 6
    Const Cst_i期別 As Integer = 7
    Const Cst_i上課地址 As Integer = 8
    Const Cst_i停辦 As Integer = 9
    Const Cst_i上課時段 As Integer = 10
    Const Cst_i師資 As Integer = 11
    Const Cst_i助教 As Integer = 20  '20120213 BY AMU (產投用助教)

    Const Cst_i核定人數 As Integer = 12  'Cst_招生人數  as Integer=12
    Const Cst_i增班 As Integer = 13
    Const Cst_i科場地 As Integer = 14 '學(術)科場地
    Const Cst_i上課時間 As Integer = 15
    Const Cst_i其他 As Integer = 16
    Const Cst_i報名日期 As Integer = 17  '20080825 andy  add 報名日期
    Const Cst_i課程表 As Integer = 18  '20080626 andy add 課程表

    Const Cst_i包班種類 As Integer = 19  '20111208 BY AMU 
    Const Cst_i訓練費用 As Integer = 21  '20170908 (職前)
    Const Cst_i遠距教學 As Integer = 22  '2021/06/09'增修需求 OJT-21060201 產投 - 班級變更申請/審核：新增遠距教學變更 + 網站-顯示遠距教學資訊 DISTANCE learning /distance teaching
    Const Cst_iMaxChgItem As Integer = 22

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
        '檢查Session是否存在 End
        Call TIMS.OpenDbConn(objconn)

        PageControler1.PageDataGrid = DataGrid1
        'iPYNum14=TIMS.sUtl_GetPYNum14(Me)

        '產學訓套用的顯示字串/'非產學訓套用的顯示字串
        ChgItemName = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, TIMS.TPlanID28ChgItemName, TIMS.TPlanIDChgItemName)

        'gda.SelectCommand.CommandText=sql
        gsCmd = Get_SEL_REVISE_SQLCMD1(objconn)

        If Not IsPostBack Then
            Call CCreate1()
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

    Public Shared Function Get_SEL_REVISE_SQLCMD1(ByRef oConn As SqlConnection) As SqlCommand
        'gda設定 'gda=TIMS.GetOneDA(objconn)
        Dim sql As String = ""
        sql &= " SELECT ISNULL(max(PTDRID),0) PTDRID FROM PLAN_TRAINDESC_REVISE" & vbCrLf
        sql &= " WHERE PlanID=@planid AND ComIDNO=@comidno AND SeqNO=@seqno AND CDate=@cdate AND SubSeqNO=@subseqno" & vbCrLf
        Return New SqlCommand(sql, oConn)
    End Function

    ''' <summary> 變更項目設計 </summary>
    Sub SHOW_CHGITEM_1()

        '調整可變更內容-- --Start
        Dim ChgItemSortVal As String = "1,2,3,4,5,6,7,8,9,10,11,20,12,13,14,15,18,22,17,19,21,16"
        Dim ChgItemName As String()
        '**by Milor 20080507--將變更項目的顯示字串，使用陣列管理，如果需要依不同條件套不同名稱的話，可以直接在這邊修改----start
        '2008-05-21 andy 新增「課程表」
        'ChgItemName=New String() {"開、結訓日期", "訓練時段", "訓練課程地點", "課程編配", "訓練師資", "班別", "期別", "上課地址", "停辦", "上課時段", "師資", "核定人數", "增班", "上課地點", "上課時間", "其他", "報名日期", "課程表", "包班種類"}
        'ChgItemName=New String() {"訓練期間", "訓練時段", "訓練課程地點", "課程編配", "訓練師資", "班別名稱", "期別", "上課地址", "申請停辦", "上課時段", "師資", "核定人數", "增班", "學(術)科場地", "上課時間", "其他", "報名日期", "包班種類"}
        'DISTANCE '遠距教學 BY AMU 20210610

        '產學訓套用的顯示字串/'非產學訓套用的顯示字串
        ChgItemName = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, TIMS.TPlanID28ChgItemName, TIMS.TPlanIDChgItemName)

        Dim ChgItemSortAry As String() = ChgItemSortVal.Split(",")
        With ChgItem.Items
            .Clear() '清理
            .Add(New ListItem(TIMS.cst_ddl_PleaseChoose3, "")) '請選擇
            For i_CI As Integer = 0 To ChgItemSortAry.Length - 1
                Dim str_val As String = ChgItemSortAry(i_CI)
                Dim str_txt As String = ChgItemName(Val(str_val) - 1)
                .Add(New ListItem(str_txt, str_val))
            Next
        End With

        'Cst_i包班種類
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If flag_PackageType_NOUSE Then 'true:(未使用)未選擇 包班種類(移除) 'false:(使用)選擇 包班種類(保留)
                '未選擇 包班種類
                ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i包班種類)) '未選擇 包班種類
            End If
        End If

        '產投使用／遠距教學 暫不啟用
        flag_StopDISTANCE = If(TIMS.Utl_GetConfigSet("STOP_DISTANCE").Equals("Y"), True, False)

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '企訓專用／產投／充電 自辦職前
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i訓練時段))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i訓練地點))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i課程編配))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i訓練師資))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i班別名稱))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i上課時段))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i增班))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i報名日期)) '20081015 andy 報名日期產投無此項目
            '**by Milor 20080507--97年產學訓只剩下1.開、結訓日期；9.停辦；11.師資；14.上課地點；15.上課時間；16.其他----start
            'If sm.UserInfo.Years >= 2008 Then
            'End If
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i期別)) '2008停用 期別:7
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i上課地址)) '2008停用 上課地址:8
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i核定人數)) '201111 '開放因應無薪假再出發 '201203'有配套措失，手冊沒有此功能故移除
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i訓練費用)) '產投 暫無訓練費用 

            '20080722 andy edit  暫開放其它項目，因課程表變更尚未上線
            'ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i其他))
            'ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i包班種類))
            '**by Milor 20080507----end
            '產業人才投資方案的上課時間／時段問題
            '2008/1/24 由 上課時間 改為 上課時段 '2008/4/24再改回 上課時間 以後將不再改為上課時段 by 豪哥/AMU
            'ChgItem.Items.FindByValue("15").Text="上課時段"
            '產投使用／遠距教學 暫不啟用
            If flag_StopDISTANCE Then ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i遠距教學))
        Else
            '一般TIMS計畫 ，非產投類
            '未開班把不能變更狀態的移除
            'If ViewState(vs_OCID)="" Then
            '    ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i訓練地點))
            '    ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i訓練師資))
            'End If
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i上課時段))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i核定人數))
            '經PM發現 TIMS計劃應該無此功能(停辦:9、師資:11、增班:13)-  2008-01-03 by AMU
            'ChgItem.Items.Remove(ChgItem.Items.FindByValue("9"))  '開放停辦功能
            '經PM發現一般TIMS計劃應該無此功能(停辦)-  2008-01-03 by  Andy FindByValue("9"))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i停辦)) ''停辦
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i師資))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i助教)) '職前移除
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i增班)) ''增班
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i科場地))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i上課時間))
            'ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i其他)) '20080923 一般計畫其它項目保留
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i報名日期))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i課程表))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i包班種類))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i遠距教學)) '產投使用／在職暫不用
        End If
        '調整可變更內容 --End

        'Common.SetListItem(ChgItem, ChgStateVal)
    End Sub

    Sub CCreate1()
        ROC_Years.Value = (sm.UserInfo.Years - 1911)

        SHOW_CHGITEM_1()

        DataGridTable.Visible = False
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        PlanPoint = TIMS.Get_RblPlanPoint(Me, PlanPoint, objconn)
        Common.SetListItem(PlanPoint, "1")

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))
        'Years.Value=sm.UserInfo.Years - 1911
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Call sSearch1()
    End Sub

    Sub sSearch1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        Dim v_ChgItem As String = TIMS.GetListValue(ChgItem)
        Dim v_CheckMode As String = TIMS.GetListValue(CheckMode)

        Dim parms As Hashtable = New Hashtable()
        parms.Add("TPlanID", sm.UserInfo.TPlanID)
        parms.Add("Years", sm.UserInfo.Years)
        'parms.Clear()

        Dim sql As String = ""
        sql &= " SELECT a.PlanID, a.ComIDNO, a.SeqNO ,c.ORGNAME" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType) ClassName" & vbCrLf
        sql &= " ,format(b.CDate,'yyyy/MM/dd') CDate" & vbCrLf
        sql &= " ,b.SubSeqNO,b.ALTDATAID" & vbCrLf
        sql &= " ,b.ReviseStatus,b.PARTREDUC" & vbCrLf
        sql &= " ,a.TMID ,a.RID ,c.RELSHIP ,c.OrgKind,c.DISTID" & vbCrLf
        sql &= " FROM PLAN_PLANINFO a" & vbCrLf
        sql &= " JOIN PLAN_REVISE b ON b.PlanID=a.PlanID AND b.ComIDNO=a.ComIDNO AND b.SeqNO=a.SeqNO" & vbCrLf
        sql &= " JOIN VIEW_RIDNAME c ON a.RID=c.RID" & vbCrLf
        sql &= " LEFT JOIN CLASS_CLASSINFO cc ON cc.PlanID=a.PlanID AND cc.ComIDNO=a.ComIDNO AND cc.SeqNO=a.SeqNO" & vbCrLf
        sql &= " WHERE c.TPlanID=@TPlanID AND c.Years=@Years" & vbCrLf
        If sm.UserInfo.LID <> 0 Then
            sql &= " AND a.PlanID=@PlanID" & vbCrLf
            parms.Add("PlanID", sm.UserInfo.PlanID)
        End If
        '職類/班別
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        If OCIDValue1.Value <> "" Then
            sql &= " AND cc.OCID=@OCID AND a.RID=@RID" & vbCrLf
            parms.Add("OCID", OCIDValue1.Value)
            parms.Add("RID", RIDValue.Value)
        Else
            Dim sRelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)
            If sRelShip <> "" Then
                sql &= " AND c.RelShip LIKE @RelShip+'%'" & vbCrLf
                parms.Add("RelShip", sRelShip)
            Else
                sql &= " AND a.RID=@RID" & vbCrLf
                parms.Add("RID", RIDValue.Value)
            End If
        End If
        '班級名稱
        If ClassName.Text <> "" Then
            sql &= " AND a.ClassName LIKE @ClassName" & vbCrLf
            parms.Add("ClassName", "%" & Replace(ClassName.Text, " ", "%") & "%")
        End If
        '期別
        CyclType.Text = TIMS.ClearSQM(CyclType.Text)
        If CyclType.Text <> "" AndAlso IsNumeric(CyclType.Text) Then
            If CyclType.Text.Length < 2 Then CyclType.Text = "0" & CyclType.Text
            sql &= " AND a.CyclType=@CyclType" & vbCrLf
            parms.Add("CyclType", CyclType.Text)
        End If
        '變更項目
        If v_ChgItem <> "" Then
            sql &= " AND b.ALTDATAID=@ALTDATAID" & vbCrLf
            parms.Add("ALTDATAID", v_ChgItem)
        End If
        '審核狀態   '1:審核不通過/2:審核中/0:審核完成
        Select Case v_CheckMode
            Case "1"
                sql &= " AND b.ReviseStatus='N'" & vbCrLf
            Case "2"
                sql &= " AND b.ReviseStatus IS NULL" & vbCrLf
            Case "0"
                sql &= " AND b.ReviseStatus='Y'" & vbCrLf
        End Select
        '28:產業人才投資方案
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case PlanPoint.SelectedValue
                Case "1"
                    sql &= " AND c.OrgKind <> 10" & vbCrLf
                Case "2"
                    sql &= " AND c.OrgKind=10" & vbCrLf
            End Select
        End If

        Dim flag_chktest As Boolean = TIMS.sUtl_ChkTest()
        If (flag_chktest) Then
            Dim s_parms As String = TIMS.GetMyValue3(parms)
            TIMS.WriteLog(Me, "##SD_14_010 sql:" & sql)
            TIMS.WriteLog(Me, "##SD_14_010 s_parms :" & s_parms)
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        'DataGridTable.Visible=False
        'msg.Text="查無資料"
        'If TIMS.Get_SQLRecordCount(sql, objconn) > 0 Then
        '    DataGridTable.Visible=True
        '    msg.Text=""
        'End If

        msg.Text = "查無資料"
        DataGridTable.Visible = False
        divTip.Visible = False

        If dt Is Nothing Then Return
        If dt.Rows.Count = 0 Then Return

        'If dt.Rows.Count > 0 Then End If
        divTip.Visible = True
        '28:產業人才投資方案
        KindValue.Value = TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn)
        Kind2Value.Value = KindValue.Value 'TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn)
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Dim PNAME As String = PlanPoint.SelectedItem.Text
            Select Case PlanPoint.SelectedValue
                Case "1", "2"
                    KindValue.Value &= "（" & PNAME & "）"
                    Kind2Value.Value = PNAME
            End Select
        End If
        msg.Text = ""
        DataGridTable.Visible = True
        'PageControler1.SqlDataCreate(sql)
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

        'edit，by:20181101
        'divTip.Visible=If(dt.Rows.Count > 0, True, False)
    End Sub

    '空白變更申請書
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '28:產業人才投資方案
        'KindValue.Value=TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn)
        'If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    Dim PNAME As String=PlanPoint.SelectedItem.Text
        '    Select Case PlanPoint.SelectedValue
        '        Case "1", "2"
        '            KindValue.Value &= "（" & PNAME & "）"
        '    End Select
        'End If
        '#{Years}'#{RID}'#{PlanID}'#{DISTID}
        'And IP.YEARS='2019''And IP.DISTID='001''And c.RID='B5465''And c.PlanID='4823'
        'ROC_Years.Value=(sm.UserInfo.Years - 1911)
        Dim myValue As String = "GRED=" & TIMS.GetRnd6Eng()
        myValue &= "&Years=" & sm.UserInfo.Years
        myValue &= "&RID=" & sm.UserInfo.RID
        myValue &= "&PlanID=" & sm.UserInfo.PlanID
        myValue &= "&DISTID=" & sm.UserInfo.DistID
        '空白變更申請書
        prtFilename = cst_printFN3 '"SD_14_010_1_b"
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, "Years=" & sm.UserInfo.Years & "&Title=" & Convert.ToString(KindValue.Value) & "&PCName=" & Convert.ToString(Kind2Value.Value))
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, myValue)
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                'e.Item.CssClass="SD_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""
                Dim Button3 As HtmlInputButton = e.Item.FindControl("Button3")   '(列印)計畫變更表
                Dim bt_print As HtmlInputButton = e.Item.FindControl("bt_print") '(列印)變更後課程表
                Dim SCDate As String = If(Not IsDBNull(drv("CDate")), TIMS.Cdate3(drv("CDate")), "") 'yyyy/MM/dd

                'Dim SMpath As String=ReportQuery.GetSmartQueryPath
                '-20081016 andy add 列印變更課程表(產學訓) -課程表-start '停辦、其它、非產學訓
                Dim sAltDataID As String = Convert.ToString(drv("AltDataID"))
                bt_print.Disabled = (sAltDataID = "9" OrElse sAltDataID = "16" OrElse SCDate < CDate("2008-9-20"))
                If (bt_print.Disabled) Then TIMS.Tooltip(bt_print, "停辦、其它", True)
                If Not bt_print.Disabled Then
                    Dim iPTDRID As Integer = Get_PTDRID(drv("PlanID"), drv("ComIDNO"), drv("SeqNO"), SCDate, drv("SubSeqNO"), gsCmd)
                    'prtFilename=If(Convert.ToString(drv("TMID"))=cst_EHour_Use_TMID, SD_14_010.cst_printFN1d, SD_14_010.cst_printFN1c)
                    prtFilename = If(Convert.ToString(drv("TMID")) = TIMS.cst_EHour_Use_TMID, cst_printFN1d, cst_printFN1c)
                    bt_print.Attributes("onclick") = $"openPrint('../../SQControl.aspx?filename={prtFilename}&PTDRID={iPTDRID}&AltDataID={sAltDataID}');"
                End If
                '-20081016 andy add 列印變更課程表(產學訓) -課程表-end

                '列印變更申請表
                prtFilename = cst_printFN2 '"SD_14_010_b"
                'Button3.Attributes("onclick")="openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain&filename=" & prtFilename & "&path=" & SMpath & "&Years=" & (sm.UserInfo.Years - 1911) & "&PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNo") & "&CDate=" & SCDate & "&SubSeqNO=" & drv("SubSeqNO") & "&Title='+escape('" & KindValue.Value & "')+' &');"
                'http://163.29.199.222:8080/ReportServer3/report.do?GUID=38506&RptID=SD_14_010_b&Years=107&PlanID=4519&ComIDNO=40760667&SeqNo=15&CDate=2019/01/08&SubSeqNO=1&Title=x&UserID=L7100071
                'select dbo.FN_REV_PLAN_ONCLASS(4519,'40760667',15,1,convert(date,'2019/01/08'),'VWANDT')
                'Dim REVIEW_JS As String="openPrint('../../SQControl.aspx?filename=" & prtFilename & "&Years=" & (sm.UserInfo.Years - 1911) & "&PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNo=" & drv("SeqNo") & "&CDate=" & SCDate & "&SubSeqNO=" & drv("SubSeqNO") & "&Title='+escape('" & KindValue.Value & "')+'&');"
                Dim REVIEW_JS As String = $"openPrint('../../SQControl.aspx?filename={prtFilename}&Years={ROC_Years.Value}&PlanID={drv("PlanID")}&ComIDNO={drv("ComIDNO")}&SeqNo={drv("SeqNo")}&CDate={SCDate}&SubSeqNO={drv("SubSeqNO")}');"
                Button3.Attributes("onclick") = REVIEW_JS
                Dim s_AltDataID_txtN As String = ""
                Try
                    s_AltDataID_txtN = $"{ChgItemName(CInt(sAltDataID) - 1)} ({drv("SubSeqNO")})"
                Catch ex As Exception
                    s_AltDataID_txtN = "<FONT color='red'>陣列資料錯誤</FONT>"
                End Try
                e.Item.Cells(Cst_變更項目).Text = s_AltDataID_txtN

                '**by Milor 20080502--只有審核中的資料才能按列印----start
                Dim S_REVISESTATUS_TXT As String = cst_ReviseStatus_txt_oth
                Select Case $"{drv("ReviseStatus")}"
                    Case ""
                        'S_REVISESTATUS_TXT = cst_ReviseStatus_txt_oth '"審核中"
                        S_REVISESTATUS_TXT = If($"{drv("PARTREDUC")}" = TIMS.cst_YES, cst_ReviseStatus_txt_PARTREDUC_Y, cst_ReviseStatus_txt_oth)
                        Button3.Disabled = False
                        TIMS.Tooltip(Button3, S_REVISESTATUS_TXT, True)
                        'Button3.Style.Add("background-color", "lightgray")  'edit，by:20181120
                        If ($"{drv("PARTREDUC")}" = TIMS.cst_YES) Then
                            'OJT-25071402：<系統> 產投_訓練計畫變更表：調整還原時之【審核狀態】及增加列印按鈕卡控
                            '增加卡控，當審核狀態為「還原修正中」，則「訓練計畫變更表」及「變更後課程表」鈕反灰，不提供列印
                            Const cst_tit1 As String = "審核狀態為「還原修正中」，不可列印"
                            If Not Button3.Disabled Then
                                Button3.Disabled = True '(列印)計畫變更表
                                TIMS.Tooltip(Button3, cst_tit1, True)
                            End If
                            If Not bt_print.Disabled Then
                                bt_print.Disabled = True '(列印)變更後課程表
                                TIMS.Tooltip(bt_print, cst_tit1, True)
                            End If
                        End If
                    Case "Y"
                        S_REVISESTATUS_TXT = cst_ReviseStatus_txt_Y '"審核通過"
                        Button3.Disabled = True
                        Button3.Style.Add("background-color", "lightgray")
                        'edit，by:20181120  'If (Not bt_print.Disabled) Then bt_print.Disabled=True
                        TIMS.Tooltip(Button3, S_REVISESTATUS_TXT, True)
                    Case "N"
                        S_REVISESTATUS_TXT = cst_ReviseStatus_txt_N '"審核失敗"
                        Button3.Disabled = True
                        Button3.Style.Add("background-color", "lightgray")  'edit，by:20181120
                        'If (Not bt_print.Disabled) Then bt_print.Disabled=True
                        TIMS.Tooltip(Button3, S_REVISESTATUS_TXT, True)
                End Select
                e.Item.Cells(Cst_審核狀態).Text = S_REVISESTATUS_TXT
                '**by Milor 20080502----end

                If TIMS.sUtl_ChkTest() Then '測試用
                    If Button3.Disabled Then
                        Button3.Disabled = False
                        TIMS.Tooltip(Button3, "測試平台開啟功能@@")
                    End If
                End If

        End Select
    End Sub

    Public Shared Function Get_PTDRID(ByVal PLANID As Integer, ByVal COMIDNO As String, ByVal SEQNO As Integer, ByVal tmpDATE As String, ByVal tmpSubSNO As Integer, ByRef gsCmd As SqlCommand) As Integer
        Dim iRst As Integer = 0
        Try
            With gsCmd
                .Parameters.Clear()
                .Parameters.Add("planid", SqlDbType.Int).Value = PLANID
                .Parameters.Add("comidno", SqlDbType.VarChar).Value = COMIDNO
                .Parameters.Add("seqno", SqlDbType.Int).Value = SEQNO
                .Parameters.Add("cdate", SqlDbType.DateTime).Value = Convert.ToDateTime(tmpDATE)
                .Parameters.Add("subseqno", SqlDbType.Int).Value = tmpSubSNO
                iRst = .ExecuteScalar()
            End With
        Catch ex As Exception
            Dim strErrmsg As String = $"ex.Message: {ex.Message}{vbCrLf}" '取得錯誤資訊寫入
            strErrmsg &= $"planid:{PLANID},comidno:{COMIDNO},seqno:{SEQNO},cdate:{tmpDATE},subseqno:{tmpSubSNO}{vbCrLf}{TIMS.GetErrorMsg()}"
            strErrmsg = Replace(strErrmsg, vbCrLf, $"<br>{vbCrLf}")
            Call TIMS.WriteTraceLog(strErrmsg, ex)
        End Try
        Return iRst
    End Function

#Region "(No Use)"

    'Private Function getAltItemName(ByVal AltDataID As String) As String
    '    Dim AltItemName As String=""
    '    Select Case AltDataID
    '        Case "1"
    '            AltItemName="訓練期間"
    '        Case "2"
    '            AltItemName="訓練時段"
    '        Case "3"
    '            AltItemName="訓練課程地點"
    '        Case "4"
    '            AltItemName="課程編配"
    '        Case "5"
    '            AltItemName="訓練師資"
    '        Case "6"
    '            AltItemName="班別"
    '        Case "7"
    '            AltItemName="期別"
    '        Case "8"
    '            AltItemName="上課地址"
    '        Case "9"
    '            AltItemName="申請停辦"
    '        Case "10"
    '            AltItemName="上課時段"
    '        Case "11"
    '            AltItemName="師資"
    '        Case "12"
    '            AltItemName="招生人數"
    '        Case "13"
    '            AltItemName="增班"
    '        Case "14"
    '            AltItemName="學(術)科場地"
    '        Case "15"
    '            AltItemName="上課時間"
    '        Case "16"
    '            AltItemName="其他"
    '        Case "17"
    '            AltItemName="報名日期"
    '        Case "18"
    '            AltItemName="課程表"
    '        Case "19"
    '            AltItemName="包班種類"
    '        Case Else
    '            AltItemName="未設定"
    '    End Select
    '    Return AltItemName
    'End Function

    'Private Function Get_PTDRID(ByVal planid As Integer, ByVal comidno As String, ByVal seqno As Integer, ByVal adate As String, ByVal subseqno As Integer) As Integer
    '    Dim objConn As SqlConnection=DbAccess.GetConnection()
    '    Dim sqlAdp As New SqlDataAdapter
    '    Dim sqlStr As String
    '    Dim rst As Integer=0
    '    Try
    '        objConn.Open()
    '        sqlStr="select PTDRID from Plan_TrainDesc_Revise" & vbCrLf
    '        sqlStr += "where PlanID=@planid and ComIDNO=@comidno and SeqNO=@seqno and CDate=@cdate and SubSeqNO=@subseqno "
    '        With sqlAdp
    '            .SelectCommand=New SqlCommand(sqlStr, objConn)
    '            .SelectCommand.Parameters.Clear()
    '            .SelectCommand.Parameters.Add("@planid", SqlDbType.Int).Value=planid
    '            .SelectCommand.Parameters.Add("@comidno", SqlDbType.VarChar).Value=comidno
    '            .SelectCommand.Parameters.Add("@seqno", SqlDbType.Int).Value=seqno
    '            .SelectCommand.Parameters.Add("@cdate", SqlDbType.DateTime).Value=Convert.ToDateTime(adate)
    '            .SelectCommand.Parameters.Add("@subseqno", SqlDbType.Int).Value=subseqno
    '            rst=.SelectCommand.ExecuteScalar()
    '        End With
    '    Catch ex As Exception
    '        Common.MessageBox(Me, ex.ToString)
    '    Finally
    '        objConn.Close()
    '    End Try
    '    Return rst
    'End Function

#End Region
End Class