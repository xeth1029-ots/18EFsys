Partial Class SD_05_012
    Inherits AuthBasePage

    'dbo.FN_GET_CHECKCLASS
    'STUD_QUESTIONFAC2
    'STUD_QUESTIONARY,V_STUDQUESTION1
    'U  Class_ClassInfo SET CanClose=NULL,CanCloseResult=NULL,CanCloseACCT=NULL,CanCloseDATE=NULL WHERE OCID =@OCID
    '--s CanCloseResult,CanCloseACCT,CanCloseDATE from CLASS_CLASSINFO where RID in ('B6153')
    '--u CLASS_CLASSINFO  set CanCloseResult =null,CanCloseACCT=null,CanCloseDATE =null from CLASS_CLASSINFO where RID in ('B6153') and (CanCloseResult is not null or CanCloseACCT is not null or CanCloseDATE is not null )
    Dim ff As String = ""
    Dim dtArc As DataTable '暫時權限Table
    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。

    Const cst_Checkbox1 As Integer = 0        '勾選
    Const cst_seq As Integer = 1              '序號
    Const cst_TrainName As Integer = 2        '訓練職類
    Const cst_ClassID As Integer = 3          '班級代碼
    Const cst_ClassCName As Integer = 4       '班級名稱
    Const cst_STDate As Integer = 5           '開訓日期
    Const cst_FTDate As Integer = 6           '結訓日期
    Const cst_StudentCount As Integer = 7     '學員人數
    Const cst_StudentClose As Integer = 8     '結訓人數
    Const cst_nodata As Integer = 9           '未填資料
    Const cst_CanCloseResult As Integer = 10  '開放班級結訓理由

    Dim iCntRow1 As Integer = 0  '1筆資料時啟用特殊權限

    Dim objconn As SqlConnection

    'UPDATE Class_ClassInfo SET CanClose=NULL,CanCloseResult=NULL,CanCloseACCT=NULL,CanCloseDATE=NULL WHERE OCID =@OCID
    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        PageControler1.PageDataGrid = DataGrid1
        '檢查Session是否存在 End

        '暫時權限Table------------------------------Start
        ''Dim dtArc As DataTable '暫時權限Table
        dtArc = TIMS.Get_Auth_REndClass(Me, objconn)
        '暫時權限Table------------------------------End
        '是否為超級使用者
        flgROLEIDx0xLIDx0 = TIMS.IsSuperUser(Me, 1)

        '產投啟用。(有值啟用。)
        DataGrid1.Columns(cst_CanCloseResult).Visible = False
        labMsg28.Visible = False
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case Convert.ToString(sm.UserInfo.LID)
                Case "0", "1" '發展署、分署啟用該功能。(產投啟用該功能。)
                    labMsg28.Visible = True
                    HidTPlanID.Value = sm.UserInfo.TPlanID '產投啟用。
                    DataGrid1.Columns(cst_CanCloseResult).Visible = True
            End Select
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button7.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        Button2.Attributes("onclick") = "return Check_Data();"

        If Not IsPostBack Then
            msg.Text = ""
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            DataGridTable.Visible = False
            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            'TIMS.GET_OnlyOne_OCID(CStr(RIDValue.Value))
            'Button3_Click(sender, e)
        End If
    End Sub

    '搜尋其中之1欄位資料。
    Function Get_CanCloseResult(ByVal OCIDvalue As String) As String
        Dim rst As String = ""
        If OCIDvalue = "" Then Return rst
        For Each item As DataGridItem In DataGrid1.Items
            Dim Star As Label = item.FindControl("Star")
            Dim Checkbox1 As HtmlInputCheckBox = item.FindControl("Checkbox1")
            Dim ChangeFlag As HtmlInputHidden = item.FindControl("ChangeFlag") '有異動。
            Dim CanCloseResult As TextBox = item.FindControl("CanCloseResult") '開放班級結訓理由。
            If Checkbox1.Value = OCIDvalue Then
                rst = TIMS.ClearSQM(CanCloseResult.Text)
                Exit For
            End If
        Next
        Return rst
    End Function

    '開放可做結訓填寫理由
    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "OpenCmd" '開放可做結訓填寫理由
                Dim s_ocidvalue As String = TIMS.GetMyValue(e.CommandArgument, "OCID")
                Dim drCC As DataRow = TIMS.GetOCIDDate(s_ocidvalue, objconn)
                If drCC Is Nothing Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Exit Sub
                End If

                Dim s_CanCloseResult As String = Get_CanCloseResult(s_ocidvalue)
                If s_CanCloseResult = "" Then
                    Common.MessageBox(Me, "請填寫 開放可做結訓填寫理由!")
                    Exit Sub
                End If

                Dim sql As String = ""
                sql &= " UPDATE Class_ClassInfo" & vbCrLf
                sql &= " SET CanClose=@CanClose" & vbCrLf
                sql &= " ,CanCloseResult=@CanCloseResult" & vbCrLf
                sql &= " ,CanCloseACCT=@CanCloseACCT" & vbCrLf
                sql &= " ,CanCloseDATE=GETDATE()" & vbCrLf
                sql &= " WHERE OCID=@OCID" & vbCrLf
                Dim parms As New Hashtable ' parms.Clear()
                parms.Add("CanClose", "Y") '可做結訓。
                parms.Add("CanCloseResult", s_CanCloseResult)
                parms.Add("CanCloseACCT", Convert.ToString(sm.UserInfo.UserID))
                parms.Add("OCID", s_ocidvalue)
                DbAccess.ExecuteNonQuery(sql, objconn, parms)

                Call sUtl_Search1() '搜尋。[SQL]
        End Select
    End Sub

    Private Sub DataGrid1_ItemCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemCreated
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"
                Dim i As Integer = -1
                Dim MyImage As New System.Web.UI.WebControls.Image
                If Not Me.ViewState("sort") Is Nothing Then
                    Select Case Me.ViewState("sort")
                        Case "TrainID", "TrainID DESC"
                            i = 2
                        Case "ClassID", "ClassID DESC"
                            i = 3
                        Case "CLASSCNAME2", "CLASSCNAME2 DESC"
                            i = 4
                        Case "STDate", "STDate DESC"
                            i = 5
                        Case "FTDate", "FTDate DESC"
                            i = 6
                    End Select
                    MyImage.ImageUrl = If(ViewState("sort").ToString.IndexOf(" DESC") = -1, "../../images/SortUp.gif", "../../images/SortDown.gif")
                    If i <> -1 Then e.Item.Cells(i).Controls.Add(MyImage)
                End If
            Case ListItemType.Item
                e.Item.CssClass = ""
            Case ListItemType.AlternatingItem
        End Select
    End Sub

    ''' <summary>FinCheck</summary>
    ''' <param name="s_OCID"></param>
    ''' <returns></returns>
    Function GET_CHECKCLASS(ByVal s_OCID As String) As String
        Dim rst As String = ""
        If s_OCID = "" OrElse Not TIMS.IsNumeric2(s_OCID) Then Return rst
        Dim sql As String = String.Format("SELECT dbo.FN_GET_CHECKCLASS({0}) FinCheck", s_OCID)
        rst = Convert.ToString(DbAccess.ExecuteScalar(sql, objconn))
        Return rst
    End Function

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim Result As Array
        'Dim I As Integer
        Dim ResStr As String = ""
        Dim SubStr As String = ""
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1") 'Checkbox1.Value: OCID
                Dim ChangeFlag As HtmlInputHidden = e.Item.FindControl("ChangeFlag")
                Dim HidCanClose As HtmlInputHidden = e.Item.FindControl("HidCanClose")
                Dim BtnCanClose As Button = e.Item.FindControl("BtnCanClose") '開放
                BtnCanClose.CommandArgument = String.Concat("OCID=", drv("OCID")) '開放
                Dim CanCloseResult As TextBox = e.Item.FindControl("CanCloseResult") '開放可做結訓填寫理由
                Dim star As Label = e.Item.FindControl("star") '表示為該班有必填資料未填，無法執行班級結訓動
                e.Item.Cells(cst_seq).Text = TIMS.Get_DGSeqNo(sender, e) '序號
                Checkbox1.Disabled = False '未鎖定
                Checkbox1.Checked = False '未勾選
                If drv("IsClosed").ToString = "Y" Then
                    'RoleID: 0:管理者 1:系統管理者 5:承辦人
                    Select Case Convert.ToString(sm.UserInfo.RoleID)
                        Case "0", "1", "5"
                            If String.Equals(Convert.ToString(drv("StudentCount")), "0") Then
                                Checkbox1.Disabled = True '鎖定
                                TIMS.Tooltip(Checkbox1, "學員人數為0", True)
                            End If
                        Case Else
                            Checkbox1.Disabled = True '鎖定
                            TIMS.Tooltip(Checkbox1, "角色不為系統管理者或承辦人", True)
                    End Select
                    Checkbox1.Checked = True '勾選 '已經結訓。
                Else
                    If String.Equals(Convert.ToString(drv("StudentCount")), "0") Then
                        Checkbox1.Disabled = True '鎖定
                        TIMS.Tooltip(Checkbox1, "學員人數為0", True)
                    End If
                    Checkbox1.Checked = False '未勾選 未結訓。
                End If

                '說明未填寫的資訊
                ResStr = "" 'dbo.fn_GET_CheckClass
                Dim s_FinCheck As String = GET_CHECKCLASS(Convert.ToString(drv("OCID")))
                If s_FinCheck <> "X" Then
                    Result = Split(s_FinCheck, ";") '依分號切割。
                    For iI As Integer = 0 To UBound(Result)
                        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                            Select Case Left(Result(iI), 1)
                                Case "D" '產投流程要確認 CREDITPOINTS
                                    '產投流程 ', "46", "47" '產投及補辦
                                    SubStr = Mid(Result(iI).ToString, 3, Len(Result(iI).ToString) - 2)
                                    If SubStr <> "0" Then ResStr &= String.Concat("結訓成績(產): ", SubStr, "人未填<BR>")
                                Case "F"
                                    '產投流程要確認 受訓學員意見調查表 STUD_QUESTIONFAC/STUD_QUESTIONFAC2
                                    SubStr = Mid(Result(iI).ToString, 3, Len(Result(iI).ToString) - 2)
                                    If SubStr <> "0" Then ResStr &= String.Concat("受訓學員意見調查表(產): ", SubStr, "人未填")
                            End Select
                        Else
                            Select Case Left(Result(iI), 1)
                                Case "D" '流程確認 CREDITPOINTS
                                    SubStr = Mid(Result(iI).ToString, 3, Len(Result(iI).ToString) - 2)
                                    If SubStr <> "0" Then ResStr &= String.Concat("核發結訓證書: ", SubStr, "人未填<BR>")
                                Case "A"
                                    '非 (產投及補辦)
                                    SubStr = Mid(Result(iI).ToString, 3, Len(Result(iI).ToString) - 2)
                                    If SubStr <> "0" Then ResStr &= String.Concat("結訓成績: ", SubStr, "人未填<BR>")
                                Case "B"
                                    SubStr = Mid(Result(iI).ToString, 3, Len(Result(iI).ToString) - 2)
                                    If SubStr <> "0" Then ResStr &= "結訓學員資料卡封面檔未填<BR>"
                                Case "C"
                                    SubStr = Mid(Result(iI).ToString, 3, Len(Result(iI).ToString) - 2)
                                    If SubStr <> "0" Then ResStr &= String.Concat("結訓學員資料卡: ", SubStr, "人未填<BR>")
                                Case "E"
                                    SubStr = Mid(Result(iI).ToString, 3, Len(Result(iI).ToString) - 2)
                                    If SubStr <> "0" Then ResStr &= String.Concat("訓練期末學員滿意度: ", SubStr, "人未填")
                            End Select
                        End If

                    Next

                End If

                '說明未填寫的資訊
                e.Item.Cells(cst_nodata).Text = ResStr
                'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '    If Convert.ToString(drv("StudQFcnt")) <> "" AndAlso CInt(Convert.ToString(drv("StudQFcnt"))) > 0 Then ResStr=ResStr + "受訓學員意見調查表(產): 尚有" & Convert.ToString(drv("StudQFcnt")) & "人未填"
                'End If
                'If ResStr="" Then e.Item.Cells(10).Text="Y" Else e.Item.Cells(10).Text="N"
                If ResStr <> "" Then '有未填寫資料
                    star.Visible = True
                    TIMS.Tooltip(star, ResStr, True)
                    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        Select Case Convert.ToString(sm.UserInfo.LID)
                            Case "0", "1" '發展署、分署啟用該功能。(產投啟用該功能。)
                                '未勾選，可再勾選 為未結訓
                                If Not Checkbox1.Checked Then '若未勾選結訓 ，則 Disabled=True
                                    '有未填寫資料
                                    Checkbox1.Disabled = True '鎖定
                                    'TIMS.Tooltip(Checkbox1, "有未填寫資料")
                                End If
                            Case Else '委訓單位
                                '有未填寫資料
                                Checkbox1.Disabled = True '鎖定
                                'TIMS.Tooltip(Checkbox1, "有未填寫資料")
                        End Select
                    Else
                        '有未填寫資料
                        Checkbox1.Disabled = True '鎖定
                        'TIMS.Tooltip(Checkbox1, "有未填寫資料")
                    End If
                    TIMS.Tooltip(Checkbox1, ResStr, True)
                End If
                If Convert.ToDateTime(drv("FTDate")) > Now Then
                    Checkbox1.Disabled = True '鎖定 (不可使用。)
                    TIMS.Tooltip(Checkbox1, "今日尚未超過結訓日期", True)
                End If

                'TIMS 計畫 (非產投)
                If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Const Cst_sStartDay As String = "2012/08/01"
                    '起動日期為 "2012/08/01"
                    If DateDiff(DateInterval.Day, CDate(Cst_sStartDay), CDate(Today)) >= 0 Then
                        '未鎖定 比對結訓日 結訓日後30日後，功能即鎖住。
                        If Not Checkbox1.Disabled Then
                            '改為已結訓不可修改 未結訓可修改為結訓。
                            If CInt(drv("xDay1")) > 30 AndAlso Checkbox1.Checked Then
                                Checkbox1.Disabled = True '鎖定 (不可使用。)
                                TIMS.Tooltip(Checkbox1, "結訓日後30日後，功能即鎖住", True)
                            End If
                        End If
                    End If
                End If

                If Checkbox1.Disabled Then
                    '授權設定該班級有設定則開放
                    If Not TIMS.ChkIsEndDate(Convert.ToString(drv("OCID")), TIMS.cst_FunID_班級結訓作業, dtArc) Then
                        Checkbox1.Disabled = False '開放
                        TIMS.Tooltip(Checkbox1, "授權設定該班級有開放")
                    End If
                End If

                If Checkbox1.Disabled AndAlso Convert.ToString(drv("IsClosed")) = "N" AndAlso Convert.ToString(drv("CanClose")) = "Y" AndAlso Convert.ToString(drv("CanCloseResult")) <> "" Then
                    '產投。(鎖定且未結訓且開放結訓。)
                    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        Checkbox1.Disabled = False '開放
                        TIMS.Tooltip(Checkbox1, "該班級有開放結訓理由")
                    End If
                End If

                If Checkbox1.Disabled Then
                    '啟用特殊權限，可強制結訓作業請謹慎!!!
                    If flgROLEIDx0xLIDx0 AndAlso iCntRow1 = 1 Then
                        Checkbox1.Disabled = False '開放
                        TIMS.Tooltip(Checkbox1, "使用者" & LCase(sm.UserInfo.UserID) & "，可強制結訓作業請謹慎!!!")
                    ElseIf flgROLEIDx0xLIDx0 AndAlso iCntRow1 <> 1 Then
                        TIMS.Tooltip(Checkbox1, "搜尋結果為1筆資料時，強制結訓作業謹慎!!")
                    End If
                End If

                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    HidCanClose.Value = Convert.ToString(drv("CanClose"))
                    If BtnCanClose IsNot Nothing Then
                        '未做過開放理由。
                        BtnCanClose.Enabled = True
                        CanCloseResult.Enabled = True
                        If Convert.ToString(drv("IsClosed")) = "Y" Then
                            BtnCanClose.Enabled = False
                            CanCloseResult.Enabled = False
                            TIMS.Tooltip(BtnCanClose, "已經結訓")
                            TIMS.Tooltip(CanCloseResult, "已經結訓")
                        Else
                            Select Case Convert.ToString(sm.UserInfo.LID)
                                Case "0", "1" '發展署、分署啟用該功能。(產投啟用該功能。)
                                    BtnCanClose.Visible = True  '產投啟用。
                                    TIMS.Tooltip(BtnCanClose, "未完成結訓，署、分署啟用開放結訓!")
                                Case Else '委訓單位(看不到。)
                                    BtnCanClose.Visible = False  '產投啟用。
                            End Select
                            CanCloseResult.Visible = BtnCanClose.Visible  '產投啟用。
                        End If

                        '有做過開放理由。
                        If Convert.ToString(drv("CanClose")) = "Y" Then
                            BtnCanClose.Enabled = False
                            CanCloseResult.Enabled = False
                            TIMS.Tooltip(BtnCanClose, "已做過開放理由")
                            TIMS.Tooltip(CanCloseResult, "已做過開放理由")
                        End If
                        If Not CanCloseResult.Visible Then '不管是否顯示都把理由貼上。
                            e.Item.Cells(cst_CanCloseResult).Text = Convert.ToString(drv("CanCloseResult"))
                        Else
                            CanCloseResult.Text = Convert.ToString(drv("CanCloseResult"))
                        End If
                        If BtnCanClose.Enabled Then BtnCanClose.Attributes("onclick") = "return confirm('這樣將會開放結訓\n確定要開放結訓?');"
                    End If
                End If
                Checkbox1.Value = Convert.ToString(drv("OCID"))
                Checkbox1.Attributes("onclick") = "document.getElementById('" & ChangeFlag.ClientID & "').value='1';"
        End Select
    End Sub

    Private Sub DataGrid1_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid1.SortCommand
        If e Is Nothing Then Return
        Dim s_sort1 As String = TIMS.ClearSQM(Me.ViewState("sort"))
        If s_sort1 = "" Then Return
        Dim s_sort2 As String = TIMS.ClearSQM(e.SortExpression)
        If s_sort2 = "" Then Return
        ViewState("sort") = String.Concat(e.SortExpression, If(String.Equals(s_sort1, s_sort2), " DESC", ""))
        PageControler1.ChangeSort(Me.ViewState("sort"))
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call sUtl_Search1() '搜尋。[SQL]
    End Sub

    '搜尋。[SQL]
    Sub sUtl_Search1()

        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        '依sm.UserInfo.PlanID取得PlanKind  '1:自辦(內訓) 2:委外
        Dim PlanKind As String = TIMS.Get_PlanKind(Me, objconn)

        Dim sql As String = ""
        sql &= " SELECT a.OCID ,a.STDate ,a.FTDate" & vbCrLf
        '結訓日後30日後，功能即鎖住。比對結訓日與系統日期。 xDay1
        sql &= "  ,DATEDIFF(DAY ,CONVERT(DATE ,a.FTDate) ,CONVERT(DATE ,GETDATE())) xDay1" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSCNAME2" & vbCrLf
        sql &= " ,a.CyclType ,a.IsClosed" & vbCrLf
        sql &= " ,a.Years + '0' + ISNULL(b.CLASSID2,b.CLASSID) + isnull(a.CyclType,'') ClassID" & vbCrLf
        sql &= " ,ISNULL(b.CLASSID2,b.CLASSID) CLASSID1" & vbCrLf
        sql &= " ,ISNULL(d.StudentCount,0) StudentCount" & vbCrLf
        sql &= " ,ISNULL(d.StudentClose,0) StudentClose" & vbCrLf
        'sql += ",ISNULL(d.StudQFcnt,0) StudQFcnt" & vbCrLf
        'sql &= " ,dbo.FN_GET_CHECKCLASS(a.OCID) FinCheck" & vbCrLf
        sql &= " ,ISNULL(c.TrainID,c.JobID) TrainID" & vbCrLf
        sql &= " ,CASE WHEN c.TrainID IS NULL THEN '[' + c.JobID + ']' + c.JobName ELSE '[' + c.TrainID + ']' + c.TrainName END TrainName" & vbCrLf
        sql &= " ,a.CanClose ,a.CanCloseACCT ,a.CanCloseDATE ,a.CanCloseRESULT" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO a" & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.PLANID=a.PLANID" & vbCrLf
        sql &= " JOIN ID_CLASS b ON a.CLSID=b.CLSID" & vbCrLf
        sql &= " JOIN KEY_TRAINTYPE c ON a.TMID=c.TMID" & vbCrLf

        'sql="SELECT a.OCID,a.STDate,a.FTDate,a.ClassCName,a.CyclType,a.IsClosed,a.Years+'0'+b.ClassID+a.CyclType as ClassID,ISNULL(d.StudentCount,0) as StudentCount" & vbCrLf
        'sql += " ,ISNULL(d.StudentClose,0) as StudentClose,dbo.dbo.fn_GET_CheckClass(a.OCID) as FinCheck" & vbCrLf
        'sql += " ,case when c.TrainID is null then c.JobID	else c.TrainID end TrainID" & vbCrLf
        'sql += " ,case when c.TrainID is null then '['+c.JobID+']'+c.JobName  else '['+c.TrainID+']'+c.TrainName end TrainName" & vbCrLf
        'sql += " FROM" & vbCrLf
        'sql += "JOIN ID_Class b ON a.CLSID=b.CLSID" & vbCrLf
        'sql += "JOIN Key_TrainType c ON a.TMID=c.TMID" & vbCrLf

        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '假如不是產學訓計畫
            sql &= " LEFT JOIN (" & vbCrLf
            sql &= "   SELECT cc.OCID ,COUNT(1) StudentCount" & vbCrLf
            sql &= "    ,COUNT(CASE WHEN cs.StudStatus=5 THEN 1 END) StudentClose" & vbCrLf
            'sql += "   ,SUM(CASE WHEN cs.StudStatus=5 AND qf.SOCID IS NULL THEN 1 ELSE 0 END) StudQFcnt" & vbCrLf
            sql &= "   FROM CLASS_STUDENTSOFCLASS cs" & vbCrLf
            sql &= "   JOIN CLASS_CLASSINFO cc ON cc.ocid=cs.ocid" & vbCrLf
            sql &= "   JOIN ID_PLAN ip ON ip.planid=cc.planid" & vbCrLf
            'sql += "  LEFT JOIN STUD_QUESTIONFAC qf ON qf.SOCID=cs.SOCID" & vbCrLf
            sql &= "   WHERE 1=1 "
            If sm.UserInfo.RID = "A" Then
                sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
                sql &= " AND ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
            Else
                sql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            End If
            sql &= "  GROUP BY cc.OCID" & vbCrLf
            sql &= " ) d ON a.OCID=d.OCID" & vbCrLf
        Else
            '產學訓計畫,多一個判斷"是否得到學分"欄位等於1
            sql &= " LEFT JOIN (" & vbCrLf
            sql &= "   SELECT cc.OCID ,COUNT(1) StudentCount" & vbCrLf
            '排除離退訓學員輸入資料 by AMU 20090916
            sql &= "   ,COUNT(CASE WHEN cs.StudStatus=5 AND cs.CreditPoints=1 THEN 1 END) StudentClose" & vbCrLf
            '排除離退訓學員輸入 產學訓學員意見調查記錄檔(產學訓)  201108 BY AMU
            'sql += "  ,SUM(CASE WHEN cs.StudStatus=5 AND qf.SOCID IS NULL THEN 1 ELSE 0 END) StudQFcnt" & vbCrLf
            sql &= "  FROM CLASS_STUDENTSOFCLASS cs" & vbCrLf
            sql &= "  JOIN CLASS_CLASSINFO cc ON cc.ocid=cs.ocid" & vbCrLf
            sql &= "  JOIN ID_PLAN ip ON ip.planid=cc.planid" & vbCrLf
            'sql += " LEFT JOIN STUD_QUESTIONFAC qf ON qf.SOCID=cs.SOCID" & vbCrLf
            sql &= "  WHERE cs.STUDSTATUS NOT IN (2,3)" & vbCrLf
            If sm.UserInfo.RID = "A" Then
                sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
                sql &= " AND ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
            Else
                sql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            End If
            sql &= "  GROUP BY cc.OCID" & vbCrLf
            sql &= " ) d ON a.OCID=d.OCID" & vbCrLf
        End If

        sql &= " WHERE a.IsSuccess='Y' AND a.NotOpen='N'" & vbCrLf
        'sql += " AND a.PlanID='1548' AND a.RID='E1822'" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql &= " AND ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            If PlanKind = "1" Then '自辦限定處理班級。
                sql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                sql &= " AND a.OCID IN (SELECT OCID FROM Auth_AccRWClass WHERE Account='" & sm.UserInfo.UserID & "')" & vbCrLf
            Else
                sql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            End If
        End If

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If sm.UserInfo.RID = "A" Then
            '署(局)，不可太多資料
            If RIDValue.Value <> "" Then sql &= " AND a.RID='" & RIDValue.Value & "'" & vbCrLf
        Else
            '分署(中心)、委外使用
            If RIDValue.Value <> "" Then sql &= " AND a.RID LIKE '" & RIDValue.Value & "%'" & vbCrLf
        End If
        ClassID.Text = TIMS.ClearSQM(ClassID.Text)
        If ClassID.Text <> "" Then sql &= " AND ISNULL(b.CLASSID2,b.CLASSID)='" & ClassID.Text & "'" & vbCrLf
        ClassCName.Text = TIMS.ClearSQM(ClassCName.Text)
        If Trim(ClassCName.Text) <> "" Then sql &= " AND a.ClassCName LIKE '%' + '" & ClassCName.Text & "' + '%'" & vbCrLf
        CyclType.Text = TIMS.ClearSQM(CyclType.Text)
        If CyclType.Text <> "" Then
            If IsNumeric(CyclType.Text) Then
                If Int(CyclType.Text) < 10 Then
                    sql &= " AND a.CyclType='0" & Int(CyclType.Text) & "'" & vbCrLf
                Else
                    sql &= " AND a.CyclType='" & Int(CyclType.Text) & "'" & vbCrLf
                End If
            End If
        End If

        '班級範圍
        Select Case ClassRound.SelectedIndex
            Case 0 '已結訓
                'sql += " AND a.FTDate+1 <= GETDATE()" & vbCrLf
                'sql &= " AND dbo.TRUNC_DATETIME(a.FTDate) <= dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
                sql &= "  AND a.FTDate <= CONVERT(date,GETDATE())" & vbCrLf
            Case 1 '未結訓
                'sql += " AND a.FTDate+1 > GETDATE()" & vbCrLf
                'sql &= " AND dbo.TRUNC_DATETIME(a.FTDate) > dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
                sql &= "  AND a.FTDate > CONVERT(date,GETDATE())" & vbCrLf
        End Select

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        iCntRow1 = 0
        msg.Text = "查無資料"
        DataGridTable.Visible = False

        If TIMS.dtNODATA(dt) Then Return

        iCntRow1 = dt.Rows.Count
        msg.Text = ""
        DataGridTable.Visible = True

        'PageControler1.SqlPrimaryKeyDataCreate(sql, "OCID", "ClassID,CyclType")
        PageControler1.PageDataTable = dt
        PageControler1.PrimaryKey = "OCID"
        PageControler1.Sort = "ClassID,CyclType"
        PageControler1.ControlerLoad()
    End Sub

    Function check_data1(ByRef a_OCID As ArrayList, ByRef a_IsClose As ArrayList, ByRef OCIDStr As String) As String
        Dim s_msg1 As String = ""
        'Dim a_OCID As New ArrayList
        'Dim a_IsClose As New ArrayList
        Dim StrDenyClassAll As String = "" '您所選班級尚有資料未填不予儲存
        'Dim OCIDStr As String=""
        For Each item As DataGridItem In DataGrid1.Items
            Dim Star As Label = item.FindControl("Star")
            Dim Checkbox1 As HtmlInputCheckBox = item.FindControl("Checkbox1")
            Dim ChangeFlag As HtmlInputHidden = item.FindControl("ChangeFlag") '有異動。
            Dim HidCanClose As HtmlInputHidden = item.FindControl("HidCanClose") '開放結訓
            'Dim OpenResult As TextBox=item.FindControl("OpenResult") '開放班級結訓理由。
            If ChangeFlag.Value = "1" Then
                If Star.Visible = False Then
                    '沒有錯誤
                    'If Not OpenResult Is Nothing Then AOpenResult.Add(OpenResult.Text) '開放班級結訓理由。
                    Dim t_IsClose As String = "N" '未結訓(開放班級結訓)
                    If Checkbox1.Checked Then t_IsClose = "Y" '結訓

                    a_OCID.Add(Checkbox1.Value) '加1班。
                    a_IsClose.Add(t_IsClose) '未結訓(開放班級結訓) '加1班。
                    OCIDStr &= String.Concat(If(OCIDStr <> "", ",", ""), Checkbox1.Value)

                Else
                    '有錯誤(未勾選)
                    If Not Checkbox1.Checked Then
                        'If Not OpenResult Is Nothing Then AOpenResult.Add(OpenResult.Text) '開放班級結訓理由。
                        a_OCID.Add(Checkbox1.Value) '加1班。
                        a_IsClose.Add("N") '未結訓 '加1班。
                        OCIDStr &= String.Concat(If(OCIDStr <> "", ",", ""), Checkbox1.Value)

                    Else
                        '開放結訓為空，表示不開放結訓，做判斷。
                        'Dim fg_CanClose As Boolean=(HidCanClose.Value="Y")
                        If HidCanClose.Value <> "Y" Then
                            If flgROLEIDx0xLIDx0 Then
                                '(特殊權限, 強制結訓)
                                a_OCID.Add(Checkbox1.Value) '加1班。
                                a_IsClose.Add("Y") '同意結訓 '加1班。
                                OCIDStr &= String.Concat(If(OCIDStr <> "", ",", ""), Checkbox1.Value)
                            Else
                                '您所選班級尚有資料未填不予儲存
                                StrDenyClassAll &= item.Cells(cst_ClassCName).Text & vbCrLf ' + "<br>"
                            End If
                        ElseIf HidCanClose.Value = "Y" Then
                            a_OCID.Add(Checkbox1.Value) '加1班。
                            a_IsClose.Add("Y") '同意結訓 '加1班。
                            OCIDStr &= String.Concat(If(OCIDStr <> "", ",", ""), Checkbox1.Value)
                        End If

                    End If
                End If
            End If
        Next

        If OCIDStr = "" Then
            If StrDenyClassAll <> "" Then
                StrDenyClassAll &= String.Concat("", vbCrLf, "您所選班級尚有資料未填不予儲存： ", vbCrLf)
                s_msg1 &= StrDenyClassAll
                'Common.MessageBox(Me, StrDeny)
            Else
                s_msg1 &= "因您未點選班級資料無法儲存!!" & vbCrLf
                'Common.MessageBox(Me, "因您未點選班級資料無法儲存!!")
            End If
            'Exit Sub
        End If
        Return s_msg1
    End Function

    Function SAVE_DATA1(ByRef a_OCID As ArrayList, ByRef a_IsClose As ArrayList, ByRef OCIDStr As String) As String
        Dim s_msg1 As String = ""

        'Call TIMS.OpenDbConn(tConn)
        Dim sql As String = ""
        Dim tConn As SqlConnection = DbAccess.GetConnection()
        Dim trans As SqlTransaction = DbAccess.BeginTrans(tConn)
        Try
            Dim da As SqlDataAdapter = Nothing
            Dim dt As DataTable = Nothing
            Dim FTDate As New ArrayList
            sql = " SELECT * FROM CLASS_CLASSINFO WHERE OCID IN (" & OCIDStr & ") "
            dt = DbAccess.GetDataTable(sql, da, trans)
            For i As Integer = 0 To a_OCID.Count - 1 'OCID LOOP
                ff = "OCID='" & a_OCID(i) & "'"
                If dt.Select(ff).Length > 0 Then
                    Dim dr As DataRow = dt.Select(ff)(0)
                    dr("IsClosed") = a_IsClose(i) 'Y/N
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                    FTDate.Add(dr("FTDate"))
                End If
            Next
            DbAccess.UpdateDataTable(dt, da, trans)

            For i As Integer = 0 To a_OCID.Count - 1 'OCID LOOP
                sql = " SELECT * FROM CLASS_STUDENTSOFCLASS WHERE OCID IN (" & a_OCID(i) & ") "
                dt = DbAccess.GetDataTable(sql, da, trans)
                ff = "OCID='" & a_OCID(i) & "' and StudStatus IN (1,4,5)"
                For Each dr As DataRow In dt.Select(ff) 'Class_StudentsOfClass
                    dr("CloseDate") = CDate(FTDate(i))
                    dr("StudStatus") = If(a_IsClose(i) = "Y", "5", "1")
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                    'Dim dr1 As DataRow=Nothing
                    'Dim da1 As SqlDataAdapter=Nothing
                    'Dim dt1 As DataTable=Nothing
                    'sql=" SELECT * FROM Adp_TRNData WHERE SOCID='" & dr("SOCID") & "' "
                    'dt1=DbAccess.GetDataTable(sql, da1, trans)
                    'If dt1.Rows.Count <> 0 Then
                    '    dr1=dt1.Rows(0)
                    '    dr1("ARVL_STATE")=If(IsClose(i)="Y", 1, 9)
                    '    dr1("TIMSModifyDate")=Now
                    '    DbAccess.UpdateDataTable(dt1, da1, trans)
                    'End If
                    'sql=" SELECT * FROM Adp_DGTRNData WHERE SOCID='" & dr("SOCID") & "' "
                    'dt1=DbAccess.GetDataTable(sql, da1, trans)
                    'If dt1.Rows.Count <> 0 Then
                    '    dr1=dt1.Rows(0)
                    '    dr1("ARVL_STATE")=If(IsClose(i)="Y", 1, 9)
                    '    dr1("TIMSModifyDate")=Now
                    '    DbAccess.UpdateDataTable(dt1, da1, trans)
                    'End If
                    'sql=" SELECT * FROM ADP_GOVTRNDATA WHERE SOCID='" & dr("SOCID") & "' "
                    'dt1=DbAccess.GetDataTable(sql, da1, trans)
                    'If dt1.Rows.Count <> 0 Then
                    '    dr1=dt1.Rows(0)
                    '    dr1("ARVL_STATE")=If(IsClose(i)="Y", 1, 9)
                    '    dr1("TIMSModifyDate")=Now
                    '    DbAccess.UpdateDataTable(dt1, da1, trans)
                    'End If
                Next
                DbAccess.UpdateDataTable(dt, da, trans)
            Next

            '假如有填寫結訓學員資料卡封面(, 則更新人數 - ----Start)
            'sql="UPDATE Stud_DataLid SET ResultCount=a.FinCount FROM (SELECT OCID,Sum(case StudStatus when 5 then 1 else 0 end) as FinCount FROM Class_StudentsOfClass WHERE OCID IN (" & OCIDStr & ") Group By OCID) a WHERE a.OCID=Stud_DataLid.OCID" '★
            'sql=" UPDATE Stud_DataLid a SET ResultCount=(SELECT Sum(case StudStatus when 5 then 1 else 0 end) as FinCount  FROM Class_StudentsOfClass b  WHERE a.OCID=b.OCID ) WHERE a.OCID in (" & OCIDStr & ")"
            'DbAccess.ExecuteNonQuery(sql, trans)
            '假如有填寫結訓學員資料卡封面, 則更新人數 - ----End
            'DbAccess.CommitTrans(trans)

            '假如有填寫結訓學員資料卡封面(, 則更新人數 )
            'Call TIMS.OpenDbConn(tConn)
            'Dim cmd As SqlCommand
            'Dim cmd2 As SqlCommand

            sql = ""
            sql &= " SELECT ISNULL(COUNT(CASE WHEN b.StudStatus=5 THEN 1 END),0) FinCount" & vbCrLf
            sql &= " FROM CLASS_STUDENTSOFCLASS b" & vbCrLf
            sql &= " WHERE b.OCID=@OCID" & vbCrLf
            Dim sCmd2 As New SqlCommand(sql, tConn, trans)

            sql = ""
            sql &= " SELECT COUNT(1) CT1 FROM STUD_DATALID" & vbCrLf
            sql &= " WHERE OCID=@OCID" & vbCrLf
            Dim sCmd3 As New SqlCommand(sql, tConn, trans)

            sql = ""
            sql &= " UPDATE STUD_DATALID" & vbCrLf
            sql &= " SET ResultCount=@ResultCount" & vbCrLf
            sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
            sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
            sql &= " WHERE OCID=@OCID" & vbCrLf
            Dim uCmd2 As New SqlCommand(sql, tConn, trans)

            For i As Integer = 0 To a_OCID.Count - 1
                If Convert.ToString(a_OCID(i)) <> "" AndAlso Convert.ToString(a_OCID(i)) <> "0" Then
                    Dim iResultCount As Integer = 0
                    Dim iLIDCT1 As Integer = 0
                    iResultCount = 0
                    iLIDCT1 = 0
                    With sCmd2
                        .Parameters.Clear()
                        .Parameters.Add("OCID", SqlDbType.Int).Value = a_OCID(i)
                        iResultCount = .ExecuteScalar() '取得結訓數。
                    End With
                    With sCmd3
                        .Parameters.Clear()
                        .Parameters.Add("OCID", SqlDbType.Int).Value = a_OCID(i)
                        iLIDCT1 = .ExecuteScalar() '取得班級封面檔。
                    End With
                    If iLIDCT1 > 0 Then
                        With uCmd2 'UPDATE 結訓數。
                            .Parameters.Clear()
                            .Parameters.Add("ResultCount", SqlDbType.Int).Value = iResultCount
                            .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                            .Parameters.Add("OCID", SqlDbType.Int).Value = a_OCID(i)
                            .ExecuteNonQuery()
                            'DbAccess.ExecuteNonQuery(uCmd2.CommandText, trans, uCmd2.Parameters)  'edit，by:20181101
                        End With
                    End If

                End If
            Next
            DbAccess.CommitTrans(trans)

        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(Me, ex, strErrmsg)

            DbAccess.RollbackTrans(trans)
            Call TIMS.CloseDbConn(tConn)
            s_msg1 = "網路連線異常，請重新點選操作!"
            Return s_msg1
            'Common.MessageBox(Me, "網路連線異常，請重新點選操作!")
            'Exit Sub
            'Throw ex
        End Try
        Call TIMS.CloseDbConn(tConn)
        Return s_msg1
    End Function

    ''' <summary>儲存按鈕</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim iMilSec As Integer = 1000 * 1 '1秒
        Threading.Thread.Sleep(iMilSec) '假設處理某段程序需花費1毫秒 (避免機器不同步)

        Dim a_OCID As New ArrayList
        Dim a_IsClose As New ArrayList
        'Dim StrDeny As String="" '您所選班級尚有資料未填不予儲存
        Dim OCIDStr As String = ""

        Dim s_msg1 As String = check_data1(a_OCID, a_IsClose, OCIDStr)
        If s_msg1 <> "" Then
            Common.MessageBox(Me, s_msg1)
            Exit Sub
        End If
        s_msg1 = SAVE_DATA1(a_OCID, a_IsClose, OCIDStr)
        If s_msg1 <> "" Then
            Common.MessageBox(Me, s_msg1)
            Exit Sub
        End If

        s_msg1 = ""
        s_msg1 &= "班級結訓完成" & vbCrLf
        's_msg1 &= "請記得於班級結訓後，完成加退保相關作業！"& vbCrLf
        'Common.MessageBox(Me, "請記得於班級結訓後，完成加退保相關作業！")
        Common.MessageBox(Me, s_msg1)

        Call sUtl_Search1() '搜尋。[SQL]
        'PageControler1.CreateData()
    End Sub

    ''' <summary> 訓練機構 班級查詢 </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'Dim dr As DataRow
        '判斷機構是否只有一個班級
        'dr=GET_OnlyOne_OCID(Me.Page)
        'RIDValue.Value=TIMS.ClearSQM(RIDValue.Value)
        'If RIDValue.Value="" Then RIDValue.Value=sm.UserInfo.RID

        ClassCName.Text = ""
        CyclType.Text = ""
        Dim dr1 As DataRow = TIMS.GET_OnlyOne_OCID(Me, RIDValue.Value, objconn)
        If dr1 Is Nothing Then Return
        If Convert.ToString(dr1("total")) = "1" Then
            '如果只有一個班級
            ClassCName.Text = Convert.ToString(dr1("CLASSCNAME"))
            CyclType.Text = Convert.ToString(dr1("CYCLTYPE"))
            ClassRound.SelectedIndex = CInt(dr1("ClassRoundIndex"))
            'Button1_Click(sender, e)
            Call sUtl_Search1() '搜尋。[SQL]
        End If
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class

