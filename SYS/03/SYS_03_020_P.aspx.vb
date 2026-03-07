Public Class SYS_03_020_P
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    '在這裡放置使用者程式碼以初始化網頁
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        'Dim sql As String
        'sql="SELECT DISTID,NAME FROM ID_DISTRICT WHERE DISTID!='000' ORDER BY DISTID"
        'dtDIST=DbAccess.GetDataTable(sql, objconn)

        If Not Page.IsPostBack Then
            labmsg.Text = ""
            Call SUtl_Cancel1()
            tbSch.Visible = True

            Call InitObj()
        End If
    End Sub

    '功能第一次載入初始化
    Sub InitObj()
        Hid_FUNID.Value = TIMS.ClearSQM(Request("FUNID"))
        'Call Search1Value()  '記錄查詢條件
        Call Search1()
    End Sub

    '取消
    Sub SUtl_Cancel1()
        tbSch.Visible = False
        tbList.Visible = False
        tbEdit.Visible = False
    End Sub

    '清除值(及狀態設定)
    Sub ClsValue()
        'Hid_FUNID.Value=""
        Hid_FRSEQ.Value = ""
        '報表名稱代號
        txRPTNAME.Text = ""
    End Sub

    '將搜尋值加入編輯資料
    Sub CopySch2Value()
        LabFUNID.Text = ssLabFUNID.Text
        LabFUNNAME.Text = ssLabFUNNAME.Text
        txRPTNAME.Text = sRPTNAME.Text
    End Sub

    '記錄查詢條件 
    'Sub Search1Value()
    '    sRPTNAME.Text=TIMS.ClearSQM(sRPTNAME.Text)
    '    ViewState("sRPTNAME")=sRPTNAME.Text
    'End Sub

    Sub Search_FUN()
        Hid_FUNID.Value = TIMS.ClearSQM(Hid_FUNID.Value)
        If Hid_FUNID.Value = "" Then Exit Sub

        Dim parms As New Hashtable From {{"FUNID", Hid_FUNID.Value}}
        Dim sql As String = "SELECT a.FUNID,a.NAME FUNNAME,a.SPAGE FROM ID_FUNCTION a WHERE a.FUNID=@FUNID"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If TIMS.dtNODATA(dt) Then Return

        Dim dr As DataRow = dt.Rows(0)
        ssLabFUNID.Text = Convert.ToString(dr("FUNID"))
        ssLabFUNNAME.Text = Convert.ToString(dr("FUNNAME"))
        Dim s_page As String = Convert.ToString(dr("SPAGE")).ToUpper().Replace(".ASPX", "")  '(清除".aspx"關鍵字，by:20181031)
        ssLabSPAGE.Text = s_page 'Convert.ToString(dr("SPAGE"))
    End Sub


    '查詢
    Sub Search1()
        Call Search_FUN()

        labmsg.Text = "查無資料"
        tbList.Visible = False
        Hid_FUNID.Value = TIMS.ClearSQM(Hid_FUNID.Value)
        If Hid_FUNID.Value = "" Then Exit Sub

        TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DataGrid1)

        Dim parms As New Hashtable From {{"FUNID", Hid_FUNID.Value}}
        Dim sql As String = "
SELECT a.FRSEQ,a.FUNID,f.NAME FUNNAME,a.RPTNAME
FROM ID_FUNPRINT a
JOIN ID_FUNCTION f on f.FUNID=a.FUNID
WHERE a.FUNID=@FUNID
ORDER BY a.RPTNAME,a.FRSEQ"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If TIMS.dtNODATA(dt) Then Return

        'labmsg.Text="查無資料"'tbList.Visible=False
        labmsg.Text = ""
        tbList.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '取得單筆資訊(編輯)
    Sub LoadData1()
        Hid_FRSEQ.Value = TIMS.ClearSQM(Hid_FRSEQ.Value)
        Hid_FUNID.Value = TIMS.ClearSQM(Hid_FUNID.Value)
        If Hid_FRSEQ.Value = "" Then Exit Sub
        If Hid_FUNID.Value = "" Then Exit Sub

        Dim parms As New Hashtable From {{"FRSEQ", Val(Hid_FRSEQ.Value)}, {"FUNID", Hid_FUNID.Value}}
        Dim sql As String = "
SELECT a.FRSEQ,a.FUNID,f.NAME FUNNAME,f.SPAGE,a.RPTNAME
FROM ID_FUNPRINT a
JOIN ID_FUNCTION f on f.FUNID=a.FUNID
WHERE a.FRSEQ=@FRSEQ AND a.FUNID=@FUNID
"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If TIMS.dtNODATA(dt) Then Return

        Dim dr As DataRow = dt.Rows(0)
        txRPTNAME.Text = Convert.ToString(dr("RPTNAME"))
        LabFUNID.Text = Convert.ToString(dr("FUNID"))
        LabFUNNAME.Text = Convert.ToString(dr("FUNNAME"))
        Dim s_page As String = Convert.ToString(dr("SPAGE")).ToUpper().Replace(".ASPX", "")  '(清除".aspx"關鍵字，by:20181031)
        LabSPAGE.Text = s_page 'Convert.ToString(dr("SPAGE"))
    End Sub

    'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        txRPTNAME.Text = TIMS.ClearSQM(txRPTNAME.Text)
        If txRPTNAME.Text = "" Then Errmsg += "請輸入 報表名稱代號" & vbCrLf
        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '儲存
    Sub SaveData1()
        txRPTNAME.Text = TIMS.ClearSQM(txRPTNAME.Text)
        Hid_FRSEQ.Value = TIMS.ClearSQM(Hid_FRSEQ.Value)
        Hid_FUNID.Value = TIMS.ClearSQM(Hid_FUNID.Value)
        'If Hid_FRSEQ.Value="" Then Exit Sub
        If Hid_FUNID.Value = "" Then Exit Sub

        Call TIMS.OpenDbConn(objconn)
        Dim aNow As Date = TIMS.GetSysDateNow(objconn)

        Dim i_sql As String = ""
        Dim u_sql As String = ""

        Dim sql As String = "
INSERT INTO ID_FUNPRINT( FRSEQ ,FUNID ,RPTNAME ,MODIFYACCT ,MODIFYDATE )
VALUES ( @FRSEQ ,@FUNID ,@RPTNAME ,@MODIFYACCT ,GETDATE() )"
        i_sql = sql

        'Dim sql As String=""
        sql = "
UPDATE ID_FUNPRINT 
SET RPTNAME=@RPTNAME,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()
WHERE FRSEQ=@FRSEQ
"
        u_sql = sql

        '新增重複判斷
        sql = " SELECT 'X' FROM ID_FUNPRINT WHERE FUNID=@FUNID  AND RPTNAME=@RPTNAME"
        Dim siCmd As New SqlCommand(sql, objconn)

        '修改重複判斷
        sql = " SELECT 'X' FROM ID_FUNPRINT WHERE FUNID=@FUNID AND RPTNAME=@RPTNAME AND FRSEQ!=@FRSEQ"
        Dim suCmd As New SqlCommand(sql, objconn)

        txRPTNAME.Text = TIMS.ClearSQM(txRPTNAME.Text)

        If Hid_FRSEQ.Value = "" Then
            '新增(檢核)
            Dim dt1 As New DataTable
            With siCmd
                .Parameters.Clear()
                .Parameters.Add("FUNID", SqlDbType.VarChar).Value = Hid_FUNID.Value
                .Parameters.Add("RPTNAME", SqlDbType.VarChar).Value = txRPTNAME.Text
                dt1.Load(.ExecuteReader())
            End With
            If dt1.Rows.Count > 0 Then
                Common.MessageBox(Me, "該 資料已新增，請使用修改功能!!")
                Exit Sub
            End If
        Else
            '修改(檢核)
            Dim dt1 As New DataTable
            With suCmd
                .Parameters.Clear()
                .Parameters.Add("FUNID", SqlDbType.VarChar).Value = Hid_FUNID.Value
                .Parameters.Add("RPTNAME", SqlDbType.VarChar).Value = txRPTNAME.Text
                .Parameters.Add("FRSEQ", SqlDbType.Int).Value = Val(Hid_FRSEQ.Value) '/*PK*/
                dt1.Load(.ExecuteReader())
            End With
            If dt1.Rows.Count > 0 Then
                Common.MessageBox(Me, "該 資料已存在，請重新輸入!!")
                Exit Sub
            End If
        End If

        Dim i_rst As Integer = 0
        Dim str_saveok_msg As String = ""
        Dim iFRSEQ As Integer = 0
        If Hid_FRSEQ.Value = "" Then
            '新增
            iFRSEQ = DbAccess.GetNewId(objconn, "ID_FUNPRINT_FRSEQ_SEQ,ID_FUNPRINT,FRSEQ")
            Dim i_parms As New Hashtable From {
                {"FRSEQ", iFRSEQ},
                {"FUNID", Hid_FUNID.Value},
                {"RPTNAME", txRPTNAME.Text},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            i_rst = DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)

            str_saveok_msg = "新增完成!"
        Else
            iFRSEQ = Val(Hid_FRSEQ.Value)
            '修改
            Dim u_parms As New Hashtable From {
                {"RPTNAME", txRPTNAME.Text},
                {"MODIFYACCT", sm.UserInfo.UserID},
                {"FRSEQ", iFRSEQ}
            }
            i_rst = DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)

            str_saveok_msg = "修改完成!"
        End If

        If i_rst > 0 Then
            Call SUtl_Cancel1()
            tbSch.Visible = True

            Call Search1()
            If str_saveok_msg <> "" Then Common.MessageBox(Page, str_saveok_msg)
        Else
            Common.MessageBox(Page, "執行完畢，無資料更動!")
        End If
    End Sub


    '查詢鈕
    Protected Sub BtnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
        'Call Search1Value()  '記錄查詢條件
        Call Search1()
    End Sub

    '新增鈕
    Protected Sub BtnAdd1_Click(sender As Object, e As EventArgs) Handles btnAdd1.Click
        Call ClsValue()
        Call CopySch2Value()
        Call SUtl_Cancel1()
        tbEdit.Visible = True
    End Sub

    '儲存
    Protected Sub BtnSave1_Click(sender As Object, e As EventArgs) Handles btnSave1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        Call SaveData1()
    End Sub

    '回上頁
    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles btnBack1.Click
        Call SUtl_Cancel1()
        tbSch.Visible = True
        If Hid_FUNID.Value <> "" Then Call Search1()
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandName = "" Then Exit Sub
        If e.CommandArgument = "" Then Exit Sub

        Select Case e.CommandName
            Case "UPD" '修改
                Call SUtl_Cancel1()
                tbEdit.Visible = True
                Call ClsValue()
                Dim sCmdArg As String = Convert.ToString(e.CommandArgument)
                Hid_FRSEQ.Value = TIMS.GetMyValue(sCmdArg, "FRSEQ")
                Hid_FUNID.Value = TIMS.GetMyValue(sCmdArg, "FUNID")

                Call LoadData1()
                'Case "DEL" '刪除
                '    Dim sCmdArg As String=Convert.ToString(e.CommandArgument)
                '    HidHN3ID.Value=TIMS.GetMyValue(sCmdArg, "HN3ID")
                '    Call Delete1()
        End Select
    End Sub

    '表格上的元件配置
    Private Sub DataGrid1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim objDG1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'If e.Item.ItemType=ListItemType.Item Then e.Item.CssClass=""
                e.Item.Cells(0).Text = (objDG1.PageSize * objDG1.CurrentPageIndex) + e.Item.ItemIndex + 1  '序號
                Dim lbtUpdate As LinkButton = e.Item.FindControl("lbtUpdate")
                'Dim lbtDelete As LinkButton=e.Item.FindControl("lbtDelete")
                'lbtDelete.Attributes.Add("onclick", "return confirm('您確定要刪除第" & e.Item.Cells(0).Text & "筆資料嗎?');")
                Dim sCmdArg As String = ""
                Call TIMS.SetMyValue(sCmdArg, "FRSEQ", drv("FRSEQ"))
                Call TIMS.SetMyValue(sCmdArg, "FUNID", drv("FUNID"))
                'Call TIMS.SetMyValue(sCmdArg, "FUNID", drv("FUNID"))
                lbtUpdate.CommandArgument = sCmdArg
                'lbtDelete.CommandArgument=sCmdArg
        End Select
    End Sub
    Protected Sub BtnBack2_Click(sender As Object, e As EventArgs) Handles btnBack2.Click
        Dim sUrl As String = "SYS_03_020.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
        TIMS.Utl_Redirect1(Me, sUrl)
    End Sub
End Class
