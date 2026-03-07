Partial Class SD_01_009
    Inherits AuthBasePage

    Dim iallc As Integer = 0
    Dim ic1 As Integer = 0
    Dim ic2 As Integer = 0
    Dim ic3 As Integer = 0
    'Dim ic4 As Integer = 0
    'Dim iEPW As Integer = 0
    Dim iEP2P As Integer = 0

    '欄位使用。
    Const cst_序號 As Integer = 0
    Const cst_訓練單位 As Integer = 1
    Const cst_班級名稱 As Integer = 2
    Const cst_報名人數 As Integer = 3
    Const cst_網路報名 As Integer = 4
    Const cst_現場報名 As Integer = 5
    Const cst_通訊報名 As Integer = 6
    'Const cst_推介報名 As Integer = 7
    'Const cst_就服站報名 As Integer = 8
    'Const cst_一般推介單 As Integer = 7
    'Const cst_免試推介單 As Integer = 8
    Const cst_專案核定報名 As Integer = 7

    '原本的「推介人數」改為「一般推介單人數」
    '原本的「就服站報名」改為「免試推介單人數」
    '另外新增「專案核定報名人數」。
    '請將「就服單位協助報名□不區分□是□否」處，改為
    '「報名管道：□不區分□一般推介單□免試推介單□專案核定報名」

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

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            'Table4.Visible = False
            PageControler1.Visible = False
            msg.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button4_Click(sender, e)
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, Historytable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If Historytable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT oo.orgName ,cc.classCname + cc.cycltype classCname " & vbCrLf
        sql &= "  ,ISNULL(ab.chtotal,0) chtotal " & vbCrLf
        sql &= "  ,ISNULL(ab.ch1,0) ch1 " & vbCrLf
        sql &= "  ,ISNULL(ab.ch2,0) ch2 " & vbCrLf
        sql &= "  ,ISNULL(ab.ch3,0) ch3 " & vbCrLf
        sql &= "  ,ISNULL(ab.ch4,0) ch4 " & vbCrLf
        sql &= "  ,ISNULL(ab.EPW,0) EPW " & vbCrLf
        sql &= "  ,ISNULL(ab.EP2N,0) EP2N " & vbCrLf
        sql &= "  ,ISNULL(ab.EP2P,0) EP2P " & vbCrLf
        sql &= "  ,ISNULL(ab.EP2W,0) EP2W " & vbCrLf
        sql &= " FROM class_classinfo cc " & vbCrLf
        sql &= " JOIN Auth_relship ar ON ar.RID = cc.RID " & vbCrLf
        sql &= " JOIN Org_OrgInfo oo ON oo.orgid = ar.orgid " & vbCrLf
        sql &= " JOIN ID_Plan ip ON ip.planid = cc.planid " & vbCrLf
        sql &= " JOIN ( " & vbCrLf
        sql &= "  SELECT se.OCID1 " & vbCrLf
        sql &= "    ,COUNT(CASE WHEN se.enterchannel IN (1,2,3,4) THEN 1 END) chtotal " & vbCrLf
        sql &= "    ,COUNT(CASE WHEN se.enterchannel = 1 then 1 end) ch1 " & vbCrLf
        sql &= "    ,COUNT(CASE WHEN se.enterchannel = 2 then 1 end) ch2 " & vbCrLf
        sql &= "    ,COUNT(CASE WHEN se.enterchannel = 3 then 1 end) ch3 " & vbCrLf
        sql &= "    ,COUNT(CASE WHEN ISNULL(se.ENTERCHANNEL,0) = 4 AND ISNULL(se.ENTERPATH,' ') != 'W' AND ISNULL(se.ENTERPATH2,' ') != 'P' THEN 1 END) ch4 " & vbCrLf '一般推介單
        sql &= "    ,COUNT(CASE WHEN se.EnterPath = 'W' THEN 1 END) EPW " & vbCrLf '免試推介單
        'EnterPath2: N:報名登錄 /P:專案核定報名登錄 /S:特例專案核定報名登錄 
        sql &= "    ,COUNT(CASE WHEN se.ENTERPATH2 = 'N' THEN 1 END) EP2N " & vbCrLf
        sql &= "    ,COUNT(CASE WHEN se.ENTERPATH2 = 'P' THEN 1 END) EP2P " & vbCrLf '專案核定報名人數
        sql &= "    ,COUNT(CASE WHEN se.ENTERPATH2 = 'W' THEN 1 END) EP2W " & vbCrLf
        sql &= "  FROM Stud_EnterType se " & vbCrLf
        sql &= "  JOIN Stud_EnterTemp st ON st.SETID = se.SETID " & vbCrLf
        sql &= "  JOIN class_classinfo cc ON cc.ocid = se.OCID1 " & vbCrLf
        sql &= "  JOIN ID_Plan ip ON ip.planid = cc.planid " & vbCrLf
        sql &= "  WHERE 1=1" & vbCrLf
        sql &= "  AND ip.Tplanid = '" & sm.UserInfo.TPlanID & "' " & vbCrLf
        sql &= "  AND ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
        If RIDValue.Value <> "" Then sql &= " AND cc.RID LIKE '" & RIDValue.Value & "%' " & vbCrLf
        If OCIDValue1.Value <> "" Then sql &= " AND cc.OCID = '" & OCIDValue1.Value & "' " & vbCrLf
        If cjobValue.Value <> "" Then sql &= " AND cc.CJOB_UNKEY = '" & cjobValue.Value & "' " & vbCrLf
        If STDate1.Text <> "" Then sql &= " AND cc.STDATE >= " & TIMS.To_date(STDate1.Text) & vbCrLf
        If STDate2.Text <> "" Then sql &= " AND cc.STDATE <= " & TIMS.To_date(STDate2.Text) & vbCrLf
        sql &= "  GROUP BY se.OCID1 " & vbCrLf
        sql &= " ) ab ON cc.ocid = ab.ocid1 " & vbCrLf
        sql &= " where 1=1 " & vbCrLf
        sql &= "    AND ip.Tplanid = '" & sm.UserInfo.TPlanID & "' " & vbCrLf
        sql &= "    AND ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
        If RIDValue.Value <> "" Then
            sql &= " AND cc.RID LIKE '" & RIDValue.Value & "%' " & vbCrLf
        End If
        If OCIDValue1.Value <> "" Then
            sql &= " AND cc.OCID = '" & OCIDValue1.Value & "' " & vbCrLf
        End If
        If cjobValue.Value <> "" Then
            sql &= " AND cc.CJOB_UNKEY = '" & cjobValue.Value & "' " & vbCrLf
        End If
        If STDate1.Text <> "" Then
            sql &= " AND cc.STDATE >= " & TIMS.To_date(STDate1.Text) & vbCrLf
        End If
        If STDate2.Text <> "" Then
            sql &= " AND cc.STDATE <= " & TIMS.To_date(STDate2.Text) & vbCrLf
        End If
        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn)

        DataGrid1.Visible = False
        PageControler1.Visible = False
        msg.Text = "查無資料"
        msg.Visible = True
        If dt.Rows.Count > 0 Then
            DataGrid1.Visible = True
            PageControler1.Visible = True
            msg.Text = ""
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                e.Item.Cells(cst_序號).Text = TIMS.Get_DGSeqNo(sender, e) '序號
                iallc += CInt(e.Item.Cells(cst_報名人數).Text) '報名人數
                ic1 += CInt(e.Item.Cells(cst_網路報名).Text)   '網路報名
                ic2 += CInt(e.Item.Cells(cst_現場報名).Text)   '現場報名
                ic3 += CInt(e.Item.Cells(cst_通訊報名).Text)   '通訊報名
                'ic4 += CInt(e.Item.Cells(cst_一般推介單).Text)   '推介報名
                'iEPW += CInt(e.Item.Cells(cst_免試推介單).Text)   '就服站報名
                iEP2P += CInt(e.Item.Cells(cst_專案核定報名).Text)
            Case ListItemType.Footer '尾
                e.Item.Cells(cst_序號).Text = "總  計"
                e.Item.Cells(cst_報名人數).Text = iallc '報名人數
                e.Item.Cells(cst_網路報名).Text = ic1   '網路報名
                e.Item.Cells(cst_現場報名).Text = ic2   '現場報名
                e.Item.Cells(cst_通訊報名).Text = ic3   '通訊報名
                'e.Item.Cells(cst_一般推介單).Text = ic4   '推介報名
                'e.Item.Cells(cst_免試推介單).Text = iEPW '就服站報名
                e.Item.Cells(cst_專案核定報名).Text = iEP2P
                e.Item.Cells(cst_序號).ColumnSpan = 3
                e.Item.Cells(cst_訓練單位).Visible = False '前面序號使用
                e.Item.Cells(cst_班級名稱).Visible = False '前面序號使用
        End Select
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn) '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        'Table4.Visible = False
        DataGrid1.Visible = False
        PageControler1.Visible = False
        msg.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGrid1.Visible = False
        PageControler1.Visible = False
        'Table4.Visible = False
        msg.Visible = False
    End Sub
End Class