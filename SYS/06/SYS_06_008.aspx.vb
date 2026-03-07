Partial Class SYS_06_008
    Inherits AuthBasePage

    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False

    'V_ORGECFA28_OTH (ORG_ECFA28) /ORG_ECFA28_DEL
    'HyperLink2 NavigateUrl="../../Doc/ECFA28名單.zip"

    'alter TABLE [dbo].[ORG_ECFA28] 	add CATEGORY int 
    'alter TABLE [dbo].[ORG_ECFA28_DEL] 	add CATEGORY int 

    'alter TABLE [dbo].[ORG_ECFA28] 	add CONSUMABLE [varchar] (1000)  
    'alter TABLE [dbo].[ORG_ECFA28_DEL] 	add CONSUMABLE [varchar] (1000) 
    Const cst_AddUpt As Integer = 1 '新增/修改
    Const cst_Search1 As Integer = 2 '查詢
    Const cst_View As Integer = 3 '檢視
    Const cst_功能 As Integer = 10 '欄位位置

    Dim v_PERMISSION As String = ""
    Const cst_PERMISSION_Can_Use As String = "PERMISSION_Can_Use" '可使用
    Const cst_PERMISSION_No_Use As String = "PERMISSION_No_Use" '不可使用
    Const cst_PERMISSION_ACT_SCH As String = "PERMISSION_ACT_SCH"
    Const cst_PERMISSION_ACT_ADD As String = "PERMISSION_ACT_ADD"
    Const cst_PERMISSION_ACT_UPT As String = "PERMISSION_ACT_UPT"
    Const cst_PERMISSION_ACT_COPY As String = "PERMISSION_ACT_COPY"
    Const cst_PERMISSION_ACT_EXP As String = "PERMISSION_ACT_EXP"
    Const cst_PERMISSION_ACT_IMP As String = "PERMISSION_ACT_IMP"
    Const cst_PERMISSION_ACT_DEL As String = "PERMISSION_ACT_DEL"

    Const cst_View1 As String = "View1"
    Const cst_Copy1 As String = "Copy1"
    Const cst_UPT1 As String = "UPT1"
    Const cst_Del1 As String = "Del1"

    Const cst_ecfaid As String = "ecfaid"
    Const cst_SeqNo As String = "SeqNo"

    Const cst_maintainDate As String = "2011/06/07" '產業認定日
    Const cst_judgmentDate As String = "2010/12/07" '離職判斷日

    '欄位順序
    '勞保投保證號	工廠登記證號	統一編號	廠商名稱	產(行)業別	主要產品/耗用原料
    '認定類別	地址	負責人	工廠電話	員工人數	網址	適用日期
    'Const cst_aSEQNO As Integer = 0
    Const cst_aUbno As Integer = 0 '勞保投保證號 	
    Const cst_afactoryNo As Integer = 1 '工廠登記證號
    Const cst_aComIDNO As Integer = 2 '統一編號
    Const cst_aUName As Integer = 3 '廠商名稱
    Const cst_akName As Integer = 4 '產(行)業別
    Const cst_aMproduct As Integer = 5 '主要產品
    Const cst_aConsumable As Integer = 6 '耗用原料
    Const cst_aCATEGORY As Integer = 7 '認定類別
    Const cst_aAddress As Integer = 8 '地址
    Const cst_aMaster As Integer = 9 '負責人
    Const cst_aphone As Integer = 10 '工廠電話
    Const cst_aMemNum As Integer = 11 '員工人數
    Const cst_aUrl1 As Integer = 12 '網址
    'Const cst_amaintainDate As Integer = 0 '適用日期
    'Const cst_ajudgmentDate As Integer = 0 '離職判斷日
    'Const cst_aModifyAcct As Integer = 0
    'Const cst_aModifyDate As Integer = 0
    Const cst_aiMaxLength1 As Integer = 13

#Region "CHECK VALUE"
    Const cst_st字串 As String = "字串"
    Const cst_st整數 As String = "整數"
    Const cst_st字串必填 As String = "字串必填"
    Const cst_st數字必填 As String = "數字必填"
    Const cst_st日期必填 As String = "日期必填"

    ''' <summary>
    ''' 檢核輸入值3
    ''' </summary>
    ''' <param name="fN1">欄位名稱</param>
    ''' <param name="vN1">輸入內容值</param>
    ''' <param name="sType">應符合形態</param>
    ''' <returns></returns>
    Function ChkValue1(ByVal fN1 As String, ByVal vN1 As Object, ByVal sType As String) As String
        Return ChkValue1(fN1, vN1, sType, 0)
    End Function

    ''' <summary>
    ''' 檢核輸入值4
    ''' </summary>
    ''' <param name="fN1">欄位名稱</param>
    ''' <param name="vN1">輸入內容值</param>
    ''' <param name="sType">應符合形態</param>
    ''' <param name="iSize">字串長度限制</param>
    ''' <returns></returns>
    Function ChkValue1(ByVal fN1 As String, ByVal vN1 As Object, ByVal sType As String, ByVal iSize As Integer) As String
        Dim rst As String = ""
        Select Case sType
            Case cst_st字串
                If Convert.ToString(vN1) <> "" Then
                    If Convert.ToString(vN1).Length > iSize Then rst &= fN1 & "字串長度 必須小於等於" & iSize & "字數<br>"
                End If
            Case cst_st整數
                If Not TIMS.IsNumeric2(vN1) Then
                    rst &= fN1 & fN1 & " 必須為整數數字<br>"
                End If

            Case cst_st字串必填
                If Convert.ToString(vN1) <> "" Then
                    If Convert.ToString(vN1).Length > iSize Then rst &= fN1 & "字串長度 必須小於等於" & iSize & "字數<br>"
                Else
                    rst &= fN1 & " 為必填資料<br>"
                End If
            Case cst_st數字必填
                If Convert.ToString(vN1) <> "" Then
                    If Not IsNumeric(vN1) Then rst &= fN1 & "必需為數字<br>"
                Else
                    rst &= fN1 & "必須填寫<br>"
                End If
            Case cst_st日期必填
                If Convert.ToString(vN1) <> "" Then
                    If Not IsDate(vN1) Then
                        rst &= fN1 & "必須是西元年格式(yyyy/MM/dd)<br>"
                    Else
                        If CDate(vN1) < "1900/1/1" Or CDate(vN1) > "2100/1/1" Then rst &= fN1 & "範圍有誤<br>"
                    End If
                Else
                    rst &= fN1 & "必須填寫<br>"
                End If
        End Select
        Return rst
    End Function

    ''' <summary>
    ''' 檢核輸入值5
    ''' </summary>
    ''' <param name="fN1">欄位名稱</param>
    ''' <param name="vN1">輸入內容值</param>
    ''' <param name="sType">應符合形態</param>
    ''' <param name="iMin">數值大小限制-至少</param>
    ''' <param name="iMax">數值大小限制-至多</param>
    ''' <returns></returns>
    Function ChkValue1(ByVal fN1 As String, ByVal vN1 As Object, ByVal sType As String, ByVal iMin As Integer, ByVal iMax As Integer) As String
        Dim rst As String = ""
        Select Case sType
            Case cst_st整數
                If Not TIMS.IsInt(vN1) Then
                    rst &= fN1 & " 必須為整數數字<br>"
                    Return rst
                End If
                If Val(vN1) < Val(iMin) Then
                    rst &= fN1 & " 數字超過範圍<br>"
                    Return rst
                End If
                If Val(vN1) > Val(iMax) Then
                    rst &= fN1 & " 數字超過範圍<br>"
                    Return rst
                End If
        End Select
        Return rst
    End Function

#End Region

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

        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            Call sUtl_ActionType(cst_Search1)
            Call sUtl_SetSearchVal()
            Call Search1(cst_AddUpt)
        End If

    End Sub

    '查詢
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        CHK_PERMISSION(cst_PERMISSION_ACT_SCH)
        v_PERMISSION = sm.LastResultMessage
        If v_PERMISSION <> "" Then
            Common.MessageBox(Me, v_PERMISSION)
            Exit Sub
        End If

        Call sUtl_SetSearchVal()
        Call Search1(0)
    End Sub

    '查詢
    Sub Search1(ByVal i_Type As Integer)
        'Optional ByVal sType As Integer = 0
        Dim dt As DataTable
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT TOP 500 a.EcfaID" & vbCrLf
        sql &= " ,a.SeqNo,a.factoryNo" & vbCrLf
        sql &= " ,a.CATEGORY" & vbCrLf
        sql &= " ,a.kName ,a.UName" & vbCrLf
        sql &= " ,a.ComIDNO ,a.Mproduct,a.CONSUMABLE" & vbCrLf
        sql &= " ,a.OpenStatus ,a.Address" & vbCrLf
        sql &= " ,a.Master ,a.DistName" & vbCrLf
        sql &= " ,ISNULL(a.Ubno, (CASE WHEN a.bUbno IS NOT NULL THEN '*' + CONVERT(VARCHAR, a.bUbno) ELSE 'NO DATA' END)) Ubno" & vbCrLf
        sql &= " ,a.isClose ,a.maintainDate ,a.judgmentDate" & vbCrLf
        sql &= " ,a.ModifyDate ,a.phone ,a.bUbno" & vbCrLf
        sql &= " FROM V_ORGECFA28_OTH a WITH(NOLOCK)" & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        Select Case i_Type
            Case cst_AddUpt
                sql &= " AND CONVERT(date,a.ModifyDate) = CONVERT(date,GETDATE()) " & vbCrLf
        End Select
        If Me.ViewState("MDate1") <> "" Then sql &= " AND a.ModifyDate >= " & TIMS.To_date(Me.ViewState("MDate1")) & vbCrLf
        If Me.ViewState("MDate2") <> "" Then sql &= " AND a.ModifyDate <= " & TIMS.To_date(Me.ViewState("MDate2")) & vbCrLf
        If Not Me.ViewState("SeqNoEcfaID") = "" Then
            If Val(Me.ViewState("SeqNoEcfaID")) > 0 Then
                sql &= " AND (1!=1 " & vbCrLf
                sql &= " OR a.SeqNo = '" & Val(Me.ViewState("SeqNoEcfaID")) & "' " & vbCrLf  'Org_Ecfa28
                sql &= " OR a.EcfaID LIKE '%" & Me.ViewState("SeqNoEcfaID") & "%' " & vbCrLf 'Org_ECFAGrant
                sql &= " ) " & vbCrLf
            Else
                sql &= " AND a.EcfaID LIKE '%" & Me.ViewState("SeqNoEcfaID") & "%' " & vbCrLf
            End If
        End If
        If Not Me.ViewState("factoryNo") = "" Then sql &= " AND a.factoryNo LIKE '%" & Me.ViewState("factoryNo") & "%'" & vbCrLf
        If Not Me.ViewState("CATEGORY") = "" Then sql &= " AND a.CATEGORY='" & Me.ViewState("CATEGORY") & "'" & vbCrLf
        If Not Me.ViewState("kName") = "" Then sql &= " AND a.kName LIKE '%" & Me.ViewState("kName") & "%'" & vbCrLf
        If Not Me.ViewState("UName") = "" Then sql &= " AND a.UName LIKE '%" & Me.ViewState("UName") & "%'" & vbCrLf
        If Not Me.ViewState("ComIDNO") = "" Then sql &= " AND a.ComIDNO = '" & Me.ViewState("ComIDNO") & "'" & vbCrLf
        If Not Me.ViewState("Ubno") = "" Then
            sql &= " AND (1!=1 " & vbCrLf
            sql &= " OR a.Ubno LIKE '" & Me.ViewState("Ubno") & "%' " & vbCrLf  'ecfa Org_Ecfa28.Ubno
            sql &= " OR a.bUbno LIKE '" & Me.ViewState("Ubno") & "%' " & vbCrLf 'Bus_BasicData.bUbno
            sql += ") " & vbCrLf
        End If
        If Not Me.ViewState("Address") = "" Then sql &= " AND a.Address LIKE '%" & Me.ViewState("Address") & "%' " & vbCrLf
        sql &= " ORDER BY a.ModifyDate DESC " & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)

        DataGridTable1.Visible = False
        msg.Text = "查無資料!!"
        If dt.Rows.Count > 0 Then
            DataGridTable1.Visible = True
            msg.Text = ""
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    Sub sUtl_ShowList(ByVal NumValue As String, ByVal sType As String)
        Dim dt As DataTable
        Dim sql As String
        sql = ""
        sql &= " SELECT * FROM V_ORGECFA28_OTH a WHERE 1=1 "
        Select Case sType
            Case cst_ecfaid
                sql &= " AND a.ecfaid = '" & NumValue & "' "
            Case cst_SeqNo
                sql &= " AND a.SeqNo = '" & NumValue & "' "
        End Select
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then Return

        Dim dr As DataRow = Nothing
        Dim sbUbno As String = ""
        For i As Integer = 0 To dt.Rows.Count - 1
            dr = dt.Rows(i)
            If Convert.ToString(dr("bUbno")) <> "" Then
                If sbUbno <> "" Then sbUbno &= ","
                sbUbno &= Convert.ToString(dr("bUbno"))
            End If
            If i > 0 AndAlso i Mod 5 = 1 Then sbUbno &= "<br />"
        Next
        bUbno.Text = sbUbno
        maintainDate.Text = ""
        judgmentDate.Text = ""
        If dr Is Nothing Then Exit Sub

        SeqNo.Text = Convert.ToString(dr("SeqNo"))
        EcfaID.Text = Convert.ToString(dr("EcfaID"))
        factoryNo.Text = Convert.ToString(dr("factoryNo"))
        Common.SetListItem(CATEGORY, Convert.ToString(dr("CATEGORY")))
        kName.Text = Convert.ToString(dr("kName"))
        UName.Text = Convert.ToString(dr("UName"))
        ComIDNO.Text = Convert.ToString(dr("ComIDNO"))
        Mproduct.Text = Convert.ToString(dr("MProduct"))
        Consumable.Text = Convert.ToString(dr("CONSUMABLE"))
        Address.Text = Convert.ToString(dr("Address"))
        tMaster.Text = Convert.ToString(dr("Master"))
        Ubno.Text = Convert.ToString(dr("Ubno"))
        If Convert.ToString(dr("maintainDate")) <> "" Then maintainDate.Text = Common.FormatDate(dr("maintainDate"))
        If Convert.ToString(dr("judgmentDate")) <> "" Then judgmentDate.Text = Common.FormatDate(dr("judgmentDate"))
        phone.Text = Convert.ToString(dr("phone"))
        MemNum.Text = Convert.ToString(dr("MemNum"))
        Url1.Text = Convert.ToString(dr("Url1"))
        isClose.Text = Convert.ToString(dr("isClose"))
        modifyDate.Text = ""
        If Convert.ToString(dr("modifyDate")) <> "" Then modifyDate.Text = Common.FormatDate(dr("modifyDate"))
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Hid_SEQNO.Value = ""
        Dim cCmdarg As String = ""
        Select Case e.CommandName
            Case cst_Copy1
                CHK_PERMISSION(cst_PERMISSION_ACT_COPY)
                v_PERMISSION = sm.LastResultMessage
                If v_PERMISSION <> "" Then
                    Common.MessageBox(Me, v_PERMISSION)
                    Exit Sub
                End If

                Call sUtl_ActionType(cst_AddUpt)
                cCmdarg = e.CommandArgument
                Dim s_ca_ecfaid As String = TIMS.GetMyValue(cCmdarg, cst_ecfaid)
                Dim s_ca_SeqNo As String = TIMS.GetMyValue(cCmdarg, cst_SeqNo)
                If s_ca_ecfaid <> "" Then Call sUtl_ShowList(s_ca_ecfaid, cst_ecfaid)
                If s_ca_SeqNo <> "" Then Call sUtl_ShowList(s_ca_SeqNo, cst_SeqNo)
                SeqNo.Text = "[系統預設]"
                EcfaID.Text = "[不使用]"

            Case cst_View1
                CHK_PERMISSION(cst_PERMISSION_ACT_SCH)
                v_PERMISSION = sm.LastResultMessage
                If v_PERMISSION <> "" Then
                    Common.MessageBox(Me, v_PERMISSION)
                    Exit Sub
                End If

                Call sUtl_ActionType(cst_View)
                cCmdarg = e.CommandArgument
                Dim s_ca_ecfaid As String = TIMS.GetMyValue(cCmdarg, cst_ecfaid)
                Dim s_ca_SeqNo As String = TIMS.GetMyValue(cCmdarg, cst_SeqNo)
                If s_ca_ecfaid <> "" Then Call sUtl_ShowList(s_ca_ecfaid, cst_ecfaid)
                If s_ca_SeqNo <> "" Then Call sUtl_ShowList(s_ca_SeqNo, cst_SeqNo)

            Case cst_UPT1
                CHK_PERMISSION(cst_PERMISSION_ACT_UPT)
                v_PERMISSION = sm.LastResultMessage
                If v_PERMISSION <> "" Then
                    Common.MessageBox(Me, v_PERMISSION)
                    Exit Sub
                End If

                Call sUtl_ActionType(cst_AddUpt)
                cCmdarg = e.CommandArgument
                Dim s_ca_ecfaid As String = TIMS.GetMyValue(cCmdarg, cst_ecfaid)
                Dim s_ca_SeqNo As String = TIMS.GetMyValue(cCmdarg, cst_SeqNo)
                Hid_SEQNO.Value = TIMS.ClearSQM(s_ca_SeqNo)
                If s_ca_ecfaid <> "" Then Call sUtl_ShowList(s_ca_ecfaid, cst_ecfaid)
                If s_ca_SeqNo <> "" Then Call sUtl_ShowList(s_ca_SeqNo, cst_SeqNo)

            Case cst_Del1
                CHK_PERMISSION(cst_PERMISSION_ACT_DEL)
                v_PERMISSION = sm.LastResultMessage
                If v_PERMISSION <> "" Then
                    Common.MessageBox(Me, v_PERMISSION)
                    Exit Sub
                End If

                cCmdarg = e.CommandArgument
                Dim ca_SeqNo As String = TIMS.GetMyValue(cCmdarg, cst_SeqNo)
                Dim iRst As Integer = 0
                If ca_SeqNo <> "" Then iRst = sUtl_DELETE1(ca_SeqNo)
                If iRst = 0 Then
                    Common.MessageBox(Me, "查無刪除資訊!!")
                    Return
                End If
                Common.MessageBox(Me, "刪除成功!!")

        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim cCmdarg As String = ""
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btnView1 As Button = e.Item.FindControl("btnView1")
                Dim btnCopy1 As Button = e.Item.FindControl("btnCopy1")
                Dim btnUPT1 As Button = e.Item.FindControl("btnUPT1")
                Dim btnDel1 As Button = e.Item.FindControl("btnDel1")

                cCmdarg = ""
                If Convert.ToString(drv("ecfaid")) <> "" Then TIMS.SetMyValue(cCmdarg, "ecfaid", Convert.ToString(drv("ecfaid")))
                If Convert.ToString(drv("SeqNo")) <> "" Then TIMS.SetMyValue(cCmdarg, "SeqNo", Convert.ToString(drv("SeqNo")))
                btnView1.CommandArgument = cCmdarg
                btnCopy1.CommandArgument = cCmdarg
                btnUPT1.CommandArgument = cCmdarg
                btnDel1.CommandArgument = cCmdarg

                btnUPT1.Visible = False
                btnDel1.Visible = False
                If Convert.ToString(drv("SeqNo")) <> "" Then
                    btnUPT1.Visible = True '可修改
                    btnDel1.Visible = True '可刪除
                    Dim js1 As String = "return confirm('確定要刪除 流水號：" & Convert.ToString(drv("SeqNo")) & " 的資料?');"
                    btnDel1.Attributes.Add("onclick", js1)
                End If
        End Select
    End Sub

    Sub sUtl_ClearList(ByVal i_Type As Integer)
        SeqNo.Text = ""
        EcfaID.Text = ""
        SeqNo.Enabled = False
        EcfaID.Enabled = False
        maintainDate.Enabled = False
        judgmentDate.Enabled = False
        isClose.Enabled = False
        modifyDate.Enabled = False
        bUbno.Text = ""
        maintainDate.Text = "2011/06/07"
        judgmentDate.Text = "2010/12/07"
        factoryNo.Text = ""
        CATEGORY.SelectedIndex = -1
        kName.Text = ""
        UName.Text = ""
        ComIDNO.Text = ""
        Mproduct.Text = ""
        Consumable.Text = ""
        Address.Text = ""
        tMaster.Text = ""
        Ubno.Text = ""
        phone.Text = ""
        MemNum.Text = ""
        Url1.Text = ""
        isClose.Text = ""
        Select Case i_Type
            Case cst_AddUpt
                SeqNo.Text = "[系統預設]"
                EcfaID.Text = "[不使用]"
        End Select
    End Sub

    '顯示動作改變
    Sub sUtl_ActionType(ByVal i_Type As Integer)
        Call sUtl_ClearList(i_Type)
        panelSearch.Visible = False
        panelEdit.Visible = False
        btnSave1.Visible = False
        Select Case i_Type
            Case cst_View
                panelEdit.Visible = True
            Case cst_AddUpt
                panelEdit.Visible = True
                btnSave1.Visible = True
            Case cst_Search1
                panelSearch.Visible = True
                btnSave1.Visible = True
        End Select
    End Sub

    Private Sub BtnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAdd.Click
        CHK_PERMISSION(cst_PERMISSION_ACT_ADD)
        v_PERMISSION = sm.LastResultMessage
        If v_PERMISSION <> "" Then
            Common.MessageBox(Me, v_PERMISSION)
            Exit Sub
        End If

        Call sUtl_ActionType(cst_AddUpt)
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        Call sUtl_ActionType(cst_Search1)
    End Sub

    Sub SaveDate1_INSERT()
        Dim aNow As Date '*DB現在時間
        aNow = TIMS.GetSysDateNow(objconn)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " INSERT INTO ORG_ECFA28(SEQNO, factoryNo,CATEGORY ,kName ,UName ,ComIDNO ,Mproduct,CONSUMABLE ,Address ,Master ,Ubno " & vbCrLf
        sql &= " ,maintainDate ,judgmentDate ,phone ,MemNum ,Url1 ,ModifyAcct ,ModifyDate )" & vbCrLf
        sql &= " VALUES(@SEQNO, @factoryNo,@CATEGORY ,@kName ,@UName ,@ComIDNO ,@Mproduct,@CONSUMABLE ,@Address ,@Master ,@Ubno " & vbCrLf
        sql &= " ,@maintainDate ,@judgmentDate ,@phone ,@MemNum ,@Url1 ,@ModifyAcct ,GETDATE() ) " & vbCrLf

        '取得 新的SEQNO
        'Dim mySEQNO As Integer = 0
        'Dim tSql As String = " SELECT (MAX(SEQNO) + 1) mySEQ FROM ORG_ECFA28"
        'Dim tDr As DataRow = DbAccess.GetOneRow(tSql, objconn)
        'If Not tDr Is Nothing Then mySEQNO = Convert.ToInt64(tDr("mySEQ"))

        factoryNo.Text = TIMS.ClearSQM(factoryNo.Text)
        'Dim v_CATEGORY As String = TIMS.ClearSQM(CATEGORY.SelectedValue)
        Dim v_CATEGORY As String = TIMS.GetListValue(CATEGORY)
        kName.Text = TIMS.ClearSQM(kName.Text)
        UName.Text = TIMS.ClearSQM(UName.Text)
        ComIDNO.Text = TIMS.ClearSQM(ComIDNO.Text)
        Mproduct.Text = TIMS.ClearSQM(Mproduct.Text)
        Consumable.Text = TIMS.ClearSQM(Consumable.Text)
        Address.Text = TIMS.ClearSQM(Address.Text)
        tMaster.Text = TIMS.ClearSQM(tMaster.Text)
        Ubno.Text = TIMS.ClearSQM(Ubno.Text)

        phone.Text = TIMS.ClearSQM(phone.Text)
        MemNum.Text = TIMS.ClearSQM(MemNum.Text)
        Url1.Text = TIMS.ClearSQM(Url1.Text)

        '取得 新的SEQNO
        Dim iSEQNO As Integer = DbAccess.GetNewId(objconn, "ORG_ECFA28_SEQNO_SEQ,ORG_ECFA28,SEQNO")
        Dim myParam As Hashtable = New Hashtable
        'myParam.Add("SEQNO", mySEQNO) '(pk)
        myParam.Add("SEQNO", iSEQNO) '(pk)
        myParam.Add("factoryNo", If(factoryNo.Text <> "", factoryNo.Text, Convert.DBNull))
        myParam.Add("CATEGORY", If(v_CATEGORY <> "", v_CATEGORY, Convert.DBNull))
        myParam.Add("kName", If(kName.Text <> "", kName.Text, Convert.DBNull))
        myParam.Add("UName", If(UName.Text <> "", UName.Text, Convert.DBNull))
        myParam.Add("ComIDNO", If(ComIDNO.Text <> "", ComIDNO.Text, Convert.DBNull))
        myParam.Add("Mproduct", If(Mproduct.Text <> "", Mproduct.Text, Convert.DBNull))
        myParam.Add("CONSUMABLE", If(Consumable.Text <> "", Consumable.Text, Convert.DBNull))
        myParam.Add("Address", If(Address.Text <> "", Address.Text, Convert.DBNull))
        myParam.Add("Master", If(tMaster.Text <> "", tMaster.Text, Convert.DBNull))
        myParam.Add("Ubno", If(Ubno.Text <> "", Ubno.Text, Convert.DBNull))
        myParam.Add("maintainDate", CDate(cst_maintainDate))
        myParam.Add("judgmentDate", CDate(cst_judgmentDate))
        myParam.Add("phone", If(phone.Text <> "", phone.Text, Convert.DBNull))
        myParam.Add("MemNum", If(MemNum.Text <> "", Val(MemNum.Text), Convert.DBNull))
        myParam.Add("Url1", If(Url1.Text <> "", Url1.Text, Convert.DBNull))
        myParam.Add("ModifyAcct", sm.UserInfo.UserID)
        Try
            DbAccess.ExecuteNonQuery(sql, objconn, myParam)
        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
            Return
        End Try

        Call sUtl_ActionType(cst_Search1)
        Call Search1(cst_AddUpt)
    End Sub

    Sub SaveDate1_UPDATE(ByVal iSEQNO As Integer)
        Dim aNow As Date '*DB現在時間
        aNow = TIMS.GetSysDateNow(objconn)

        Dim rst As Integer = 0
        rst = INSERT_ORG_ECFA28_DEL(iSEQNO)
        If rst = 0 Then Return

        Dim uSql As String = ""
        uSql = "" & vbCrLf
        uSql &= " UPDATE ORG_ECFA28" & vbCrLf
        uSql &= " SET FACTORYNO=@FACTORYNO" & vbCrLf
        uSql &= " ,CATEGORY=@CATEGORY" & vbCrLf
        uSql &= " ,KNAME=@KNAME" & vbCrLf
        uSql &= " ,UNAME=@UNAME" & vbCrLf
        uSql &= " ,COMIDNO=@COMIDNO" & vbCrLf
        uSql &= " ,MPRODUCT=@MPRODUCT" & vbCrLf
        uSql &= " ,CONSUMABLE=@CONSUMABLE" & vbCrLf
        uSql &= " ,ADDRESS=@ADDRESS" & vbCrLf
        uSql &= " ,MASTER=@MASTER" & vbCrLf
        uSql &= " ,UBNO=@UBNO" & vbCrLf
        uSql &= " ,MAINTAINDATE=@MAINTAINDATE" & vbCrLf
        uSql &= " ,JUDGMENTDATE=@JUDGMENTDATE" & vbCrLf
        uSql &= " ,PHONE=@PHONE" & vbCrLf
        uSql &= " ,MEMNUM=@MEMNUM" & vbCrLf
        uSql &= " ,URL1=@URL1" & vbCrLf
        uSql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        uSql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        uSql &= " WHERE 1=1" & vbCrLf
        uSql &= " AND SEQNO=@SEQNO" & vbCrLf

        factoryNo.Text = TIMS.ClearSQM(factoryNo.Text)
        'Dim v_CATEGORY As String = TIMS.ClearSQM(CATEGORY.SelectedValue)
        Dim v_CATEGORY As String = TIMS.GetListValue(CATEGORY)
        kName.Text = TIMS.ClearSQM(kName.Text)
        UName.Text = TIMS.ClearSQM(UName.Text)
        ComIDNO.Text = TIMS.ClearSQM(ComIDNO.Text)
        Mproduct.Text = TIMS.ClearSQM(Mproduct.Text)
        Consumable.Text = TIMS.ClearSQM(Consumable.Text)
        Address.Text = TIMS.ClearSQM(Address.Text)
        tMaster.Text = TIMS.ClearSQM(tMaster.Text)
        Ubno.Text = TIMS.ClearSQM(Ubno.Text)

        phone.Text = TIMS.ClearSQM(phone.Text)
        MemNum.Text = TIMS.ClearSQM(MemNum.Text)
        Url1.Text = TIMS.ClearSQM(Url1.Text)

        Dim myParam As Hashtable = New Hashtable
        myParam.Add("FACTORYNO", If(factoryNo.Text <> "", factoryNo.Text, Convert.DBNull))
        myParam.Add("CATEGORY", If(v_CATEGORY <> "", v_CATEGORY, Convert.DBNull))
        myParam.Add("KNAME", If(kName.Text <> "", kName.Text, Convert.DBNull))
        myParam.Add("UNAME", If(UName.Text <> "", UName.Text, Convert.DBNull))
        myParam.Add("COMIDNO", If(ComIDNO.Text <> "", ComIDNO.Text, Convert.DBNull))
        myParam.Add("MPRODUCT", If(Mproduct.Text <> "", Mproduct.Text, Convert.DBNull))
        myParam.Add("CONSUMABLE", If(Consumable.Text <> "", Consumable.Text, Convert.DBNull))
        myParam.Add("ADDRESS", If(Address.Text <> "", Address.Text, Convert.DBNull))
        myParam.Add("MASTER", If(tMaster.Text <> "", tMaster.Text, Convert.DBNull))
        myParam.Add("UBNO", If(Ubno.Text <> "", Ubno.Text, Convert.DBNull))
        myParam.Add("MAINTAINDATE", CDate(cst_maintainDate))
        myParam.Add("JUDGMENTDATE", CDate(cst_judgmentDate))
        myParam.Add("PHONE", If(phone.Text <> "", phone.Text, Convert.DBNull))
        myParam.Add("MEMNUM", If(MemNum.Text <> "", Val(MemNum.Text), Convert.DBNull))
        myParam.Add("URL1", If(Url1.Text <> "", Url1.Text, Convert.DBNull))
        myParam.Add("MODIFYACCT", sm.UserInfo.UserID)

        myParam.Add("SEQNO", iSEQNO) '(pk)
        Try
            DbAccess.ExecuteNonQuery(uSql, objconn, myParam)
        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
            Return
        End Try

        Call sUtl_ActionType(cst_Search1)
        Call Search1(cst_AddUpt)
    End Sub

    Private Sub btnSave1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave1.Click
        If Hid_SEQNO.Value = "" Then
            CHK_PERMISSION(cst_PERMISSION_ACT_ADD)
            v_PERMISSION = sm.LastResultMessage
            If v_PERMISSION <> "" Then
                Common.MessageBox(Me, v_PERMISSION)
                Exit Sub
            End If

            '新增
            Call SaveDate1_INSERT()
        Else
            CHK_PERMISSION(cst_PERMISSION_ACT_UPT)
            v_PERMISSION = sm.LastResultMessage
            If v_PERMISSION <> "" Then
                Common.MessageBox(Me, v_PERMISSION)
                Exit Sub
            End If

            '修改
            Call SaveDate1_UPDATE(Val(Hid_SEQNO.Value))
        End If
    End Sub

    Sub sUtl_SetSearchVal()
        'ClearSQM
        EcfaID_s.Text = TIMS.ClearSQM(EcfaID_s.Text)
        factoryNo_s.Text = TIMS.ClearSQM(factoryNo_s.Text)
        'Dim v_CATEGORY_s As String = TIMS.ClearSQM(CATEGORY_s.SelectedValue)
        Dim v_CATEGORY_s As String = TIMS.GetListValue(CATEGORY_s)
        '認定類別
        Select Case v_CATEGORY_s
            Case "1", "2" '目前只能選 1.加強輔導型產業／2.可能受貿易自由化影響產業
            Case Else
                v_CATEGORY_s = ""
        End Select
        kName_s.Text = TIMS.ClearSQM(kName_s.Text)
        UName_s.Text = TIMS.ClearSQM(UName_s.Text)
        ComIDNO_s.Text = TIMS.ClearSQM(ComIDNO_s.Text)
        Ubno_s.Text = TIMS.ClearSQM(Ubno_s.Text)
        Address_s.Text = TIMS.ClearSQM(Address_s.Text)
        MDate1.Text = TIMS.Cdate3(MDate1.Text)
        MDate2.Text = TIMS.Cdate3(MDate2.Text)

        Me.ViewState("SeqNoEcfaID") = ""
        Me.ViewState("factoryNo") = ""
        Me.ViewState("CATEGORY") = ""
        Me.ViewState("kName") = ""
        Me.ViewState("UName") = ""
        Me.ViewState("ComIDNO") = ""
        Me.ViewState("Ubno") = ""
        Me.ViewState("Address") = ""
        Me.ViewState("MDate1") = ""
        Me.ViewState("MDate2") = ""
        If EcfaID_s.Text <> "" Then Me.ViewState("SeqNoEcfaID") = TIMS.ClearSQM(EcfaID_s.Text)
        If factoryNo_s.Text <> "" Then Me.ViewState("factoryNo") = TIMS.ClearSQM(factoryNo_s.Text)
        If v_CATEGORY_s <> "" Then Me.ViewState("CATEGORY") = v_CATEGORY_s
        If kName_s.Text <> "" Then Me.ViewState("kName") = TIMS.ClearSQM(kName_s.Text)
        If UName_s.Text <> "" Then Me.ViewState("UName") = TIMS.ClearSQM(UName_s.Text)
        If ComIDNO_s.Text <> "" Then Me.ViewState("ComIDNO") = TIMS.ClearSQM(ComIDNO_s.Text)
        If Ubno_s.Text <> "" Then Me.ViewState("Ubno") = TIMS.ClearSQM(Ubno_s.Text)
        If Address_s.Text <> "" Then Me.ViewState("Address") = TIMS.ClearSQM(Address_s.Text)
        If MDate1.Text <> "" Then Me.ViewState("MDate1") = TIMS.Cdate3(MDate1.Text)
        If MDate2.Text <> "" Then Me.ViewState("MDate2") = TIMS.Cdate3(MDate2.Text)
    End Sub

    Private Sub btnSearch2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch2.Click
        CHK_PERMISSION(cst_PERMISSION_ACT_SCH)
        v_PERMISSION = sm.LastResultMessage
        If v_PERMISSION <> "" Then
            Common.MessageBox(Me, v_PERMISSION)
            Exit Sub
        End If

        Call sUtl_SetSearchVal()
        Call Search1(cst_AddUpt)
    End Sub

    Function INSERT_ORG_ECFA28_DEL(ByVal iSEQNO As Integer) As Integer
        Dim rst As Integer = 0
        Dim iSql As String = ""
        '記錄log
        iSql = "" & vbCrLf
        iSql &= " INSERT INTO ORG_ECFA28_DEL (SEQNOD ,SEQNO ,FACTORYNO,CATEGORY ,KNAME ,UNAME ,COMIDNO ,MPRODUCT,CONSUMABLE ,ADDRESS ,MASTER ,UBNO " & vbCrLf
        iSql &= "  ,MAINTAINDATE ,JUDGMENTDATE ,PHONE ,MEMNUM ,URL1 ,MODIFYACCT ,MODIFYDATE ,DELETEACCT ,DELETEDATE ) " & vbCrLf
        iSql &= " SELECT @SEQNOD ,SEQNO ,FACTORYNO,CATEGORY,KNAME ,UNAME ,COMIDNO ,MPRODUCT,CONSUMABLE ,ADDRESS ,MASTER ,UBNO" & vbCrLf
        iSql &= "  ,MAINTAINDATE ,JUDGMENTDATE ,PHONE ,MEMNUM ,URL1 ,MODIFYACCT ,MODIFYDATE ,@DELETEACCT ,GETDATE()" & vbCrLf
        iSql &= " FROM ORG_ECFA28 " & vbCrLf
        iSql &= " WHERE 1=1 " & vbCrLf
        iSql &= " AND SEQNO = @SEQNO " & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim iCmd As New SqlCommand(iSql, objconn)
        Dim iSEQNO_Del As Integer = DbAccess.GetNewId(objconn, "ORG_ECFA28_DEL_SEQNOD_SEQ,ORG_ECFA28_DEL,SEQNOD")
        Dim myParam As Hashtable = New Hashtable
        myParam.Clear()
        myParam.Add("SEQNOD", iSEQNO_Del)
        myParam.Add("DELETEACCT", sm.UserInfo.UserID)
        myParam.Add("SEQNO", iSEQNO)
        rst = DbAccess.ExecuteNonQuery(iSql, objconn, myParam)
        Return rst
    End Function

    ''' <summary>
    ''' 刪除-刪除前轉移
    ''' </summary>
    ''' <param name="iSEQNO"></param>
    ''' <returns></returns>
    Function sUtl_DELETE1(ByVal iSEQNO As Integer) As Integer
        Dim rst As Integer = 0
        rst = INSERT_ORG_ECFA28_DEL(iSEQNO)
        If rst <> 0 Then
            Dim sql_del As String = ""
            sql_del = "" & vbCrLf
            sql_del += " DELETE ORG_ECFA28" & vbCrLf
            sql_del += " WHERE 1=1" & vbCrLf
            sql_del += " AND SEQNO = @SEQNO" & vbCrLf
            'Dim myParam As Hashtable = New Hashtable
            Dim myParam As Hashtable = New Hashtable
            myParam.Clear()
            myParam.Add("SEQNO", iSEQNO)
            DbAccess.ExecuteNonQuery(sql_del, objconn, myParam)
        End If
        Return rst
    End Function

    ''' <summary>
    ''' 權限查詢(1.可做/2.不可做)
    ''' </summary>
    ''' <returns></returns>
    Function CHK_PERMISSION(ByVal C_ACT_VAL As String) As String
        Dim rst As String = cst_PERMISSION_No_Use
        '階層代碼 0:署 1:中心 2:委訓
        Select Case sm.UserInfo.LID
            Case 0 '署(全部都可以使用)
                'rst = cst_PERMISSION_Can_Use
                Return cst_PERMISSION_Can_Use 'rst
            Case Else '其它人-分署(目前僅提供查詢)
                Select Case C_ACT_VAL
                    Case cst_PERMISSION_ACT_SCH '查詢可使用(其它都不可使用)
                        'rst = cst_PERMISSION_Can_Use
                        Return cst_PERMISSION_Can_Use 'rst
                    Case cst_PERMISSION_ACT_ADD
                        sm.LastResultMessage = "不可新增"
                    Case cst_PERMISSION_ACT_UPT
                        sm.LastResultMessage = "不可修改"
                    Case cst_PERMISSION_ACT_COPY
                        sm.LastResultMessage = "不可複製"
                    Case cst_PERMISSION_ACT_EXP
                        sm.LastResultMessage = "不可匯出"
                    Case cst_PERMISSION_ACT_IMP
                        sm.LastResultMessage = "不可匯入"
                    Case cst_PERMISSION_ACT_DEL
                        sm.LastResultMessage = "不可刪除"
                End Select
        End Select
        Return rst
    End Function

    'Function GetxICmd(ByRef tConn As SqlConnection) As SqlCommand
    '    Dim sql As String = ""
    '    sql = "" & vbCrLf
    '    sql &= " INSERT INTO ORG_ECFA28(SEQNO, factoryNo,CATEGORY ,kName ,UName ,ComIDNO ,Mproduct,CONSUMABLE ,Address ,Master ,Ubno " & vbCrLf
    '    sql &= " ,maintainDate ,judgmentDate ,phone ,MemNum ,Url1 ,ModifyAcct ,ModifyDate )" & vbCrLf
    '    sql &= " VALUES(@SEQNO, @factoryNo,@CATEGORY ,@kName ,@UName ,@ComIDNO ,@Mproduct,@CONSUMABLE ,@Address ,@Master ,@Ubno " & vbCrLf
    '    sql &= " ,@maintainDate ,@judgmentDate ,@phone ,@MemNum ,@Url1 ,@ModifyAcct ,GETDATE() ) " & vbCrLf
    '    Dim iCMD As New SqlCommand(sql, tConn)
    '    Return iCMD
    'End Function

    Function ChgImpData(ByVal colArray As Array) As DataRow
        Dim dr1 As DataRow = Nothing
        Dim sql As String = ""
        sql = " SELECT * FROM ORG_ECFA28 WHERE 1<>1 "
        Dim dtV As DataTable = DbAccess.GetDataTable(sql, objconn) 'SELECT * FROM ORG_ECFA28 WHERE 1<>1
        dr1 = dtV.NewRow
        'dr1("SEQNO") = TIMS.ClearSQM(colArray(cst_aSEQNO))
        dr1("UBNO") = TIMS.ClearSQM(colArray(cst_aUbno))
        dr1("factoryNo") = TIMS.ClearSQM(colArray(cst_afactoryNo))
        dr1("COMIDNO") = TIMS.ClearSQM(colArray(cst_aComIDNO))
        dr1("UNAME") = TIMS.ClearSQM(colArray(cst_aUName))
        dr1("KNAME") = TIMS.ClearSQM(colArray(cst_akName))
        dr1("MPRODUCT") = TIMS.ClearSQM(colArray(cst_aMproduct))
        dr1("CONSUMABLE") = TIMS.ClearSQM(colArray(cst_aConsumable))
        dr1("CATEGORY") = TIMS.ClearSQM(colArray(cst_aCATEGORY))
        dr1("ADDRESS") = TIMS.ClearSQM(colArray(cst_aAddress))
        dr1("MASTER") = TIMS.ClearSQM(colArray(cst_aMaster))
        dr1("MAINTAINDATE") = TIMS.Cdate2(cst_maintainDate) 'TIMS.ClearSQM(colArray(cst_afactoryNo))
        dr1("JUDGMENTDATE") = TIMS.Cdate2(cst_judgmentDate) 'TIMS.ClearSQM(colArray(cst_afactoryNo))
        dr1("PHONE") = TIMS.Cdate2(colArray(cst_aphone))
        dr1("MEMNUM") = TIMS.ClearSQM(colArray(cst_aMemNum))
        dr1("URL1") = TIMS.ClearSQM(colArray(cst_aUrl1))
        dr1("MODIFYACCT") = sm.UserInfo.UserID
        Return dr1
    End Function

    '新增iCmd
    Sub Savedata3(ByVal dr1 As DataRow)
        Dim i_Sql As String = ""
        i_Sql = "" & vbCrLf
        i_Sql &= " INSERT INTO ORG_ECFA28(SEQNO, factoryNo,CATEGORY ,kName ,UName ,ComIDNO ,Mproduct,CONSUMABLE ,Address ,Master ,Ubno " & vbCrLf
        i_Sql &= " ,maintainDate ,judgmentDate ,phone ,MemNum ,Url1 ,ModifyAcct ,ModifyDate )" & vbCrLf
        i_Sql &= " VALUES(@SEQNO, @factoryNo,@CATEGORY ,@kName ,@UName ,@ComIDNO ,@Mproduct,@CONSUMABLE ,@Address ,@Master ,@Ubno " & vbCrLf
        i_Sql &= " ,@maintainDate ,@judgmentDate ,@phone ,@MemNum ,@Url1 ,@ModifyAcct ,GETDATE() ) " & vbCrLf

        '取得 新的SEQNO
        Dim iSEQNO As Integer = DbAccess.GetNewId(objconn, "ORG_ECFA28_SEQNO_SEQ,ORG_ECFA28,SEQNO")

        Dim myParam As Hashtable = New Hashtable
        'myParam.Add("SEQNO", mySEQNO) '(pk)
        myParam.Add("SEQNO", iSEQNO) '(pk)
        myParam.Add("factoryNo", dr1("factoryNo"))
        myParam.Add("CATEGORY", dr1("CATEGORY"))
        myParam.Add("kName", dr1("kName"))
        myParam.Add("UName", dr1("UName"))
        myParam.Add("ComIDNO", dr1("ComIDNO"))
        myParam.Add("Mproduct", dr1("Mproduct"))
        myParam.Add("CONSUMABLE", dr1("CONSUMABLE"))
        myParam.Add("Address", dr1("Address"))
        myParam.Add("Master", dr1("Master"))
        myParam.Add("Ubno", dr1("Ubno"))

        myParam.Add("maintainDate", TIMS.Cdate2(cst_maintainDate))
        myParam.Add("judgmentDate", TIMS.Cdate2(cst_judgmentDate))
        myParam.Add("phone", dr1("phone"))
        myParam.Add("MemNum", dr1("MemNum"))
        myParam.Add("Url1", dr1("Url1"))
        myParam.Add("ModifyAcct", sm.UserInfo.UserID)
        DbAccess.ExecuteNonQuery(i_Sql, objconn, myParam)

    End Sub


    '檢查輸入資料
    Function CheckImportData3(ByVal colArray As Array) As String
        Dim Reason As String = ""
        If colArray.Length < cst_aiMaxLength1 Then
            'Reason += "欄位數量不正確(應該為58個欄位)<BR>"
            Reason &= "欄位對應有誤<BR>"
            Return Reason
        End If

        '類型／必填
        Reason &= ChkValue1("勞保投保證號", colArray(cst_aUbno), cst_st字串, 20)
        Reason &= ChkValue1("工廠登記證號", colArray(cst_afactoryNo), cst_st字串, 20)

        Reason &= ChkValue1("統一編號", colArray(cst_aComIDNO), cst_st字串必填, 20)
        Reason &= ChkValue1("廠商名稱", colArray(cst_aUName), cst_st字串必填, 30)
        Reason &= ChkValue1("產(行)業別", colArray(cst_akName), cst_st字串必填, 30)
        Reason &= ChkValue1("主要產品", colArray(cst_aMproduct), cst_st字串必填, 1000)
        Reason &= ChkValue1("耗用原料", colArray(cst_aConsumable), cst_st字串, 1000)
        Reason &= ChkValue1("認定類別", colArray(cst_aCATEGORY), cst_st數字必填)
        Reason &= ChkValue1("認定類別", colArray(cst_aCATEGORY), cst_st整數, 1, 2)
        Reason &= ChkValue1("地址", colArray(cst_aAddress), cst_st字串必填, 300)
        Reason &= ChkValue1("負責人", colArray(cst_aMaster), cst_st字串必填, 50)

        Reason &= ChkValue1("工廠電話", colArray(cst_aphone), cst_st字串, 100)
        Reason &= ChkValue1("員工人數", colArray(cst_aMemNum), cst_st整數)
        Reason &= ChkValue1("網址", colArray(cst_aUrl1), cst_st字串, 200)
        'Reason &= ChkValue1("適用日期", colArray(12), cst_st數字必填)

        Return Reason
    End Function

    ''' <summary>
    ''' 匯入名單 xls
    ''' </summary>
    Sub IMP_ECFA_3B()
        'HyperLink2 NavigateUrl="../../Doc/ECFA28名單.zip"
        Dim iRowIndex As Integer = 1
        Dim sReason As String = "" '儲存錯誤的原因
        Dim dtWrong As New DataTable '儲存錯誤資料的DataTable
        Dim drWrong As DataRow = Nothing

        '建立錯誤資料格式Table
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("Reason"))

        Const Cst_Wrong_show_page As String = "SYS_06_008_Wrong.aspx"
        Const Cst_firstCol_1 As String = "勞保投保證號"
        Const Cst_FileSavePath As String = "~/CP/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Const Cst_Filetype As String = "xls"
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, Cst_Filetype) Then Return

        Dim MyFileName As String = ""
        Dim MyFileType As String = ""
        Dim dt_xls As DataTable = Nothing
        If File1.Value = "" Then
            Common.MessageBox(Me, "未輸入匯入檔案位置")
            Exit Sub
        End If
        If File1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If
        '取出檔案名稱
        MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, "檔案類型錯誤!")
            Exit Sub
        End If
        MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        If MyFileType <> Cst_Filetype Then
            Common.MessageBox(Me, "檔案類型錯誤，必須為XLS檔!")
            Exit Sub
        End If

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        '(先清理)刪除Temp中的資料
        Dim v_MyFileName As String = Server.MapPath(Cst_FileSavePath & MyFileName)
        TIMS.MyFileDelete(v_MyFileName)
        '上傳檔案 'File1.PostedFile.SaveAs(v_MyFileName) '上傳檔案
        TIMS.MyFileSaveAs(Me, File1, Cst_FileSavePath, MyFileName)
        '取得內容
        Dim fullFilNm1 As String = $"{Server.MapPath(Cst_FileSavePath & MyFileName)}"
        dt_xls = TIMS.GetDataTable_XlsFile(fullFilNm1, "", sReason, Cst_firstCol_1)
        Dim v_ErrorMsg As String = sm.LastErrorMessage
        If v_ErrorMsg <> "" Then
            Common.MessageBox(Me, v_ErrorMsg)
            Exit Sub
        End If
        'IO.File.Delete(Server.MapPath(Cst_FileSavePath & MyFileName)) '刪除檔案
        '刪除檔案
        TIMS.MyFileDelete(v_MyFileName)
        If sReason <> "" Then
            sReason &= vbCrLf
            sReason &= "資料有誤，故無法匯入，請修正Excel檔案，謝謝" & vbCrLf
            Common.MessageBox(Me, sReason)
            Exit Sub
        End If
        'xls 方式 讀取寫入資料庫
        If dt_xls.Rows.Count = 0 Then '有資料
            Common.MessageBox(Me, "查無匯入資料!!")
            Exit Sub
        End If

        '有資料
        sReason = ""
        For i As Integer = 0 To dt_xls.Rows.Count - 1
            If iRowIndex <> 0 Then
                Dim colArray As Array = dt_xls.Rows(i).ItemArray
                sReason = CheckImportData3(colArray) '依據匯入檔判斷錯誤
                If sReason = "" Then
                    Dim dr1 As DataRow = ChgImpData(colArray) 'out@dr1
                    If Not dr1 Is Nothing Then Call Savedata3(dr1) '無錯誤存檔 '匯入資料
                End If

                If sReason <> "" Then
                    '錯誤資料，填入錯誤資料表
                    drWrong = dtWrong.NewRow
                    dtWrong.Rows.Add(drWrong)
                    drWrong("Index") = iRowIndex
                    drWrong("Reason") = sReason
                End If
            End If
            iRowIndex += 1
        Next

        'tConn.BeginTransaction() 'DbAccess.GetConnection
        'Call TIMS.OpenDbConn(tConn)
        'Call TIMS.CloseDbConn(tConn)
        '判斷匯出資料是否有誤
        'Dim explain, explain2 As String
        Dim explain As String = ""
        Dim explain2 As String = ""
        explain = ""
        explain += "匯入資料共" & dt_xls.Rows.Count & "筆" & vbCrLf
        explain += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
        explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf
        explain2 = ""
        explain2 += "匯入資料共" & dt_xls.Rows.Count & "筆\n"
        explain2 += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆\n"
        explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"

        If dtWrong.Rows.Count > 0 Then
            Session("MyWrongTable") = dtWrong
            Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視原因?')){window.open('" & Cst_Wrong_show_page & "','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
            Exit Sub
        End If

        If sReason <> "" Then
            Common.MessageBox(Me, explain & sReason)
            Exit Sub
        End If

        Common.MessageBox(Me, explain)
        Exit Sub
    End Sub

    ''' <summary>
    ''' 匯入名單
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_IMP_ECFA_1_Click(sender As Object, e As EventArgs) Handles BTN_IMP_ECFA_1.Click
        CHK_PERMISSION(cst_PERMISSION_ACT_IMP)
        v_PERMISSION = sm.LastResultMessage
        If v_PERMISSION <> "" Then
            Common.MessageBox(Me, v_PERMISSION)
            Exit Sub
        End If

        Call IMP_ECFA_3B()
    End Sub

    ''' <summary>
    ''' 匯出-DataGrid1
    ''' </summary>
    Sub EXP_ECFA_3B()
        'Const cst_功能 As Integer = 9
        DataGrid1.Columns(cst_功能).Visible = False
        DataGrid1.AllowPaging = False '關閉分頁功能
        DataGrid1.EnableViewState = False  '把ViewState給關了

        '查詢
        Call sUtl_SetSearchVal()
        Call Search1(0)

        Dim sFileName As String = "匯出資料.xls"
        sFileName = HttpUtility.UrlEncode(sFileName, System.Text.Encoding.UTF8)
        Response.Clear()
        Response.Buffer = True
        Response.Charset = "UTF-8" '設定字集
        Response.AppendHeader("Content-Disposition", "attachment;filename=" & sFileName)
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        Response.ContentType = "application/ms-excel;charset=utf-8"
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        '套CSS值
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, "</style>")
        DataGrid1.AllowPaging = False '關閉分頁功能
        DataGrid1.EnableViewState = False  '把ViewState給關了
        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)
        Common.RespWrite(Me, Convert.ToString(objStringWriter))
        Response.End()
        DataGrid1.Visible = False
    End Sub

    ''' <summary>
    ''' 匯出
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_EXP_ECFA_1_Click(sender As Object, e As EventArgs) Handles BTN_EXP_ECFA_1.Click
        CHK_PERMISSION(cst_PERMISSION_ACT_EXP)
        v_PERMISSION = sm.LastResultMessage
        If v_PERMISSION <> "" Then
            Common.MessageBox(Me, v_PERMISSION)
            Exit Sub
        End If

        Call EXP_ECFA_3B()
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
