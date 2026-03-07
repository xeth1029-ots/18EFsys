Partial Class SD_01_001_sch
    Inherits AuthBasePage

    '職前專用
    '#Region "參數/變數 設定"

    'https://jira.turbotech.com.tw/browse/TIMSB-1248
    '1.勾稽資料，要排除證號為09、076、075、175、176的資料
    '2.學員資料儲存時會出現阻擋訊息, 要求服務單位必填, 頁面會清掉原勾稽選取的勞保資料-->請查明, 並修改

    'BIEF: 是否為公法救助關係 (M,N,P) (biefN/中文)
    '勞保勾稽資料畫面，於投保單位後面增加一欄"是否為公法救助關係"，並判斷為公法救助關係者，則顯示"是"
    '有關公法救助只判斷就保註記部分，如果是屬於以下四種的，都屬於公法救助關係：，詳細資料可參酌附檔(資訊室提供代號表)：
    '1.M：多元就業計畫進用人員不適用就保(9204中旬增列 )
    '2.(本項代號需與資訊室查詢)就業服務擴展計畫進用人員不適用就保(9302 下旬增列 )
    '3.N：農保被保險人參加短期就業或職業訓練僅加職災不適用就保墊償(9507中旬增列)
    '4.P：公共服務擴大就業計畫進用人員不適用就保。
    'DEPTMENT:工作部門
    'SELECT DEPTMENT,COUNT(1) CNT FROM STUD_BLIGATEDATA4 GROUP BY DEPTMENT
    'ORDER BY 1
    'SELECT BIEF,COUNT(1) CNT FROM STUD_BLIGATEDATA4 GROUP BY BIEF
    'ORDER BY 1
    'Dim gTestLc As Boolean = False 'TEST測試用

    Dim ff3 As String = ""
    Dim sNoECFA_ACTNO As String = "" '不是ECFA
    Dim sOkECFA_ACTNO As String = "" '是ECFA

    Const cst_gPublicRescue As String = "M,N,P"
    Const cst_SPAGE_SD03002 As String = "SD03002"
    Const cst_SPAGE_SD01001 As String = "SD01001"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '#Region "在這裡放置使用者程式碼以初始化網頁"

        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            msg.Text = ""
            Call create()
            Button2.Attributes("onclick") = "window.close();"
        End If

    End Sub

    '查詢
    Sub create()
        '#Region "查詢"

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Dim rqCNAME As String = TIMS.ClearSQM(TIMS.HtmlDecode1(Request("CNAME")))
        Hid_CNAME.Value = rqCNAME

        Dim rqIDNO As String = TIMS.ChangeIDNO(TIMS.ClearSQM(Request("IDNO")))
        Dim rqBIRTH As String = TIMS.Cdate3(TIMS.ClearSQM(Request("BIRTH")))
        Dim rqSPAGE As String = Request("SPAGE")
        rqSPAGE = TIMS.ClearSQM(rqSPAGE)

        'gTestLc = TIMS.sUtl_ChkTest()
        'If gTestLc Then
        '    If rqIDNO = "" Then rqIDNO = "A290010686"
        '    If rqBIRTH = "" Then rqBIRTH = "1948/10/13"
        'End If

        '1:國民身分證 2:居留證 4:居留證2021
        Dim flag1 As Boolean = TIMS.CheckIDNO(rqIDNO)
        Dim flag2 As Boolean = TIMS.CheckIDNO2(rqIDNO, 2)
        Dim flag4 As Boolean = TIMS.CheckIDNO2(rqIDNO, 4)
        If Not flag1 AndAlso Not flag2 AndAlso Not flag4 Then rqIDNO = ""

        Hid_SPAGE.Value = rqSPAGE
        Hid_idno.Value = rqIDNO
        Hid_birth.Value = rqBIRTH
        If Hid_idno.Value = "" OrElse Hid_birth.Value = "" Then
            msg.Text = "查無資料"
            DataGrid1.Visible = False
            Button1.Enabled = False
            Exit Sub
        End If
        labIDNO.Text = Convert.ToString(rqIDNO)

        Dim pms_1 As New Hashtable From {{"IDNO", rqIDNO}}
        Dim dr As DataRow = Nothing
        Dim sql As String = " SELECT b.Name FROM STUD_STUDENTINFO b WHERE b.IDNO=@IDNO"
        dr = DbAccess.GetOneRow(sql, objconn, pms_1)
        If dr IsNot Nothing Then labNAME.Text = Convert.ToString(dr("Name"))

        'Dim sql As String = ""
        'sql = "" & vbCrLf
        'sql &= " SELECT a.*" & vbCrLf
        'sql &= " FROM ADP_INTRTBLL a" & vbCrLf
        'sql &= " JOIN (" & vbCrLf
        'sql &= " 	SELECT a1.INTR_UNKEY,MAX(a1.REPLY_WKDATE) REPLY_WKDATE" & vbCrLf
        'sql &= " 	FROM ADP_INTRTBLL a1" & vbCrLf
        'sql &= " 	JOIN STUD_STUDENTINFO b1 ON a1.IDNO = b1.IDNO" & vbCrLf
        'sql &= " 	WHERE b1.IDNO = '" & rqIDNO & "'" & vbCrLf
        'sql &= " 	GROUP BY a1.INTR_UNKEY" & vbCrLf
        'sql &= " ) b ON a.INTR_UNKEY = b.INTR_UNKEY" & vbCrLf
        'dr = DbAccess.GetOneRow(sql, objconn)
        'Dim sWorkState As String = "○已就業 ●未開介紹卡"
        'If Not dr Is Nothing Then sWorkState = "●已就業 ○未開介紹卡"
        'If Not dr Is Nothing Then
        '    WorkState.Text = sWorkState '"●已就業 ○未開介紹卡"
        '    COMPNAME.Text = dr("COMPNAME").ToString
        '    COMPADDR.Text = dr("COMPADDR").ToString
        '    CONTEL.Text = dr("CONTEL").ToString
        '    REPLY_SALARY.Text = dr("REPLY_SALARY").ToString
        '    REPLY_WKDATE.Text = dr("REPLY_WKDATE").ToString
        'End If

        '查詢勞保+就保勾稽資料
        '"A"'查詢勞保+就保勾稽資料
        '"V"'查詢農保勾稽資料
        Dim s_ERRMSG As String = ""
        Try
            Call TIMS.Get_STUDBLIGATEDATA4_sch1(Me, rqIDNO, rqBIRTH, rqCNAME, s_ERRMSG, objconn)
        Catch ex As Exception
            'Common.MessageBox(Me, "勾稽時產生錯誤!!")
            Dim strErrmsg As String = ""
            strErrmsg &= "SD_01_001_sch.create():" & vbCrLf
            strErrmsg &= String.Format("ex.Message:{0}", ex.Message) & vbCrLf
            strErrmsg &= String.Format("ex.ToString:{0}", ex.ToString) & vbCrLf
            strErrmsg &= "rqIDNO:" & rqIDNO & vbCrLf
            strErrmsg &= "rqBIRTH:" & rqBIRTH & vbCrLf
            strErrmsg &= "rqCNAME:" & rqCNAME & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            Call TIMS.WriteTraceLog(strErrmsg)
            Call TIMS.CloseWin(Me, "勾稽時產生錯誤!!")
            Exit Sub
        End Try
        If s_ERRMSG <> "" Then
            Call TIMS.CloseWin(Me, s_ERRMSG)
            Exit Sub
        End If

        sql = "" & vbCrLf
        sql &= " SELECT c.SB4ID" & vbCrLf
        sql &= " ,c.IDNO" & vbCrLf
        sql &= " ,c.FTYPE" & vbCrLf
        sql &= " ,c.NAME" & vbCrLf
        sql &= " ,c.BIRTHDAY" & vbCrLf
        sql &= " ,c.ACTNO" & vbCrLf
        sql &= " ,c.COMNAME" & vbCrLf
        sql &= " ,c.CHANGEMODE" & vbCrLf
        sql &= " ,CONVERT(varchar, c.MDATE, 111) MDATE" & vbCrLf
        sql &= " ,CASE WHEN c.CHANGEMODE=4 THEN CONVERT(varchar, c.MDATE, 111) END SMDATE" & vbCrLf
        sql &= " ,CASE WHEN c.CHANGEMODE=2 THEN CONVERT(varchar, c.MDATE, 111) END FMDATE" & vbCrLf
        sql &= " ,c.SALARY" & vbCrLf
        sql &= " ,c.DEPARTMENT" & vbCrLf
        sql &= " ,c.MODIFYDATE" & vbCrLf
        sql &= " ,c.DEPTMENT" & vbCrLf
        sql &= " ,c.BIEF ,c.BIEFDESC" & vbCrLf
        sql &= " ,ISNULL(c.ComName, e.UName) UNAME" & vbCrLf
        sql &= " ,convert(varchar, getdate(), 111) TODAY1" & vbCrLf
        'sql += " ,dbo.DECODE8(c.deptment,'M','M:多元就業計畫進用人員不適用就保','N','N:農保被保險人參加短期就業或職業訓練僅加職災不適用就保墊償','P','P:公共服務擴大就業計畫進用人員不適用就保', '') deptmentN" & vbCrLf
        'sql &= " ,dbo.DECODE8(c.BIEF,'M','M:多元就業計畫進用人員不適用就保','N','N:農保被保險人參加短期就業或職業訓練僅加職災不適用就保墊償','P','P:公共服務擴大就業計畫進用人員不適用就保', '') biefN" & vbCrLf
        sql &= " ,CASE c.BIEF WHEN 'M' THEN 'M:多元就業計畫進用人員不適用就保'" & vbCrLf
        sql &= " WHEN 'N' THEN 'N:農保被保險人參加短期就業或職業訓練僅加職災不適用就保墊償'" & vbCrLf
        sql &= " WHEN 'P' THEN 'P:公共服務擴大就業計畫進用人員不適用就保' ELSE '' END biefN" & vbCrLf
        'https://jira.turbotech.com.tw/browse/TIMSC-134
        sql &= " ,case when dbo.SUBSTR(c.ACTNO,0,2) IN (" & TIMS.cst_Actno_NG2 & ") then 'Y'" & vbCrLf
        sql &= " when dbo.SUBSTR(c.ACTNO,0,3) IN (" & TIMS.cst_Actno_NG3 & ") then 'Y' END NOUSE" & vbCrLf
        sql &= " FROM dbo.STUD_BLIGATEDATA4 c" & vbCrLf
        sql &= " LEFT JOIN dbo.BUS_BASICDATA e ON c.ActNo = e.Ubno" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        '1.勾稽資料，要排除證號為09、076、075、175、176的資料
        'Public Const cst_Actno_NG2 As String = "'09'"
        'Public Const cst_Actno_NG3 As String = "'075','175','076','176'"
        'sql &= " AND dbo.SUBSTR(c.ACTNO,0,2)!='09'" & vbCrLf
        'sql &= " AND dbo.SUBSTR(c.ACTNO,0,3) NOT IN ('076','075','175','176')" & vbCrLf
        'sql &= " AND dbo.SUBSTR(c.ACTNO,0,2) NOT IN (" & TIMS.cst_Actno_NG2 & ")" & vbCrLf
        'sql &= " AND dbo.SUBSTR(c.ACTNO,0,3) NOT IN (" & TIMS.cst_Actno_NG3 & ")" & vbCrLf
        sql &= " AND c.CHANGEMODE IN (2,4)" & vbCrLf
        sql &= " AND c.IDNO = '" & rqIDNO & "'" & vbCrLf
        sql &= " AND c.BIRTHDAY = " & TIMS.To_date(rqBIRTH) & vbCrLf
        sql &= " ORDER BY c.MDATE DESC,c.CHANGEMODE DESC" & vbCrLf

        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn)
        For Each drr1 As DataRow In dt.Rows
            If Convert.ToString(drr1("SMDATE")) = "" Then
                ff3 = " ACTNO='" & Convert.ToString(drr1("ACTNO")) & "' "
                ff3 &= " AND SMDATE IS NOT NULL "
                ff3 &= " AND SMDATE <= '" & TIMS.Cdate3(drr1("FMDATE")) & "' "
                If dt.Select(ff3).Length > 0 Then drr1("SMDATE") = dt.Select(ff3, "SMDATE DESC")(0)("SMDATE")
            End If
            If Convert.ToString(drr1("FMDATE")) = "" Then
                ff3 = " ACTNO = '" & Convert.ToString(drr1("ACTNO")) & "' "
                ff3 &= " AND FMDATE IS NOT NULL "
                ff3 &= " AND FMDATE >= '" & TIMS.Cdate3(drr1("SMDATE")) & "' "
                If dt.Select(ff3).Length > 0 Then drr1("FMDATE") = dt.Select(ff3, "FMDATE DESC")(0)("FMDATE")
            End If
        Next
        dt.AcceptChanges()

        Dim sMemo As String = String.Concat("&ACT=勞保明細查詢", "&IDNO=", rqIDNO, "&BIRTH=", rqBIRTH, "&CNAME=", rqCNAME, "&sql=", sql)
        '寫入Log查詢 SubInsAccountLog1 (Auth_Accountlog)
        'https://jira.turbotech.com.tw/browse/TIMSB-1254
        'Dim s_FUNID As String = TIMS.Get_MRqID(Me)
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, "2", "", sMemo, objconn)

        Hid_ECFA_YES.Value = ""
        If dt.Rows.Count = 0 Then
            msg.Text = "查無資料"
            DataGrid1.Visible = False
            Button1.Enabled = False
            Exit Sub
        End If

        Dim dr1 As DataRow = dt.Rows(0)
        labNAME.Text = Convert.ToString(dr1("name"))

        msg.Text = ""
        DataGrid1.Visible = True
        Button1.Enabled = True

        DataGrid1.DataSource = dt
        DataGrid1.DataBind()

        Dim vMsg As String = ""
        If Hid_ECFA_YES.Value = TIMS.cst_YES Then
            vMsg = "查該民眾具有ECFA身分, 請優先選擇以ECFA身分參訓。"
            Common.MessageBox(Me, vMsg)
        End If


    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Radio1 As HtmlInputRadioButton = e.Item.FindControl("Radio1")
                Dim Hid_sb4id As HiddenField = e.Item.FindControl("Hid_sb4id")
                Dim Hid_SMDATE As HiddenField = e.Item.FindControl("Hid_SMDATE")
                Dim Hid_FMDATE As HiddenField = e.Item.FindControl("Hid_FMDATE")
                Dim LabActNoType As Label = e.Item.FindControl("LabActNoType")
                Dim LabChangeMode As Label = e.Item.FindControl("LabChangeMode")
                Dim LabECFA As Label = e.Item.FindControl("LabECFA")

                If Convert.ToString(drv("NOUSE")) = "Y" Then
                    '不可使用
                    Radio1.Disabled = True
                    TIMS.Tooltip(Radio1, "不可被點選")
                Else
                    '可使用
                    Radio1.Attributes("onclick") = "checkRadio(" & e.Item.ItemIndex + 1 & ");"
                    Radio1.Value = drv("SB4ID")
                    Hid_sb4id.Value = Convert.ToString(drv("SB4ID"))
                    Hid_SMDATE.Value = Convert.ToString(drv("SMDATE"))
                    Hid_FMDATE.Value = Convert.ToString(drv("FMDATE"))
                End If

                Dim flag_subEcfa As Boolean = False '不是ECFA
                If sNoECFA_ACTNO.IndexOf(Convert.ToString(drv("ACTNO"))) = -1 _
                    AndAlso sOkECFA_ACTNO.IndexOf(Convert.ToString(drv("ACTNO"))) = -1 Then
                    LabECFA.Text = ""
                    If TIMS.CheckIsECFA(Me, Convert.ToString(drv("ACTNO")), "", Convert.ToString(drv("TODAY1")), objconn) = True Then
                        flag_subEcfa = True '是ECFA
                        LabECFA.Text = "是"
                        Hid_ECFA_YES.Value = TIMS.cst_YES
                    End If
                End If
                If sOkECFA_ACTNO.IndexOf(Convert.ToString(drv("ACTNO"))) > -1 Then
                    flag_subEcfa = True '是ECFA
                    LabECFA.Text = "是"
                    Hid_ECFA_YES.Value = TIMS.cst_YES
                End If
                If flag_subEcfa Then
                    '是ECFA
                    If sOkECFA_ACTNO.IndexOf(Convert.ToString(drv("ACTNO"))) = -1 Then
                        '是ECFA 不須要再檢核
                        If sOkECFA_ACTNO <> "" Then sOkECFA_ACTNO &= ","
                        sOkECFA_ACTNO &= Convert.ToString(drv("ACTNO"))
                    End If
                Else
                    '不是ECFA
                    If sNoECFA_ACTNO.IndexOf(Convert.ToString(drv("ACTNO"))) = -1 Then
                        '不是ECFA 不須要再檢核
                        If sNoECFA_ACTNO <> "" Then sNoECFA_ACTNO &= ","
                        sNoECFA_ACTNO &= Convert.ToString(drv("ACTNO"))
                    End If
                End If
                'If drv("SB4ID").ToString = HidSBID.Value Then
                '    Radio1.Checked = True
                'End If

                e.Item.Cells(1).Text = e.Item.ItemIndex + 1

                Dim sActNoType As String = TIMS.Get_ACTNOTYPE1(Convert.ToString(drv("ActNo")))
                LabActNoType.Text = sActNoType

                Dim sChangeMode As String = TIMS.Get_CHANGEMODE1(Convert.ToString(drv("ChangeMode")))
                LabChangeMode.Text = sChangeMode

        End Select

    End Sub

    '#Region "勞保勾稽"

    Dim iCOUNT As Integer = 0 '取得有效資料筆數
    Dim ErrorMsg As String = "" '錯誤暫存

    'Dim sGUID As String = "" '序號(System.Guid)
    'Dim cmdSYSNAME As String = "TRA001" 'TRA001 '訓練 'TRA002 '生活津貼 FOR001'外勞
    'Dim sIDNO As String = "" 'IDNO(投保人身分證號)
    'Dim sINAME As String = "" 'NAME(投保人姓名)  
    'Dim sBIRTH As String = "" 'BIRTH(出生年月日(YYYYMMDD))
    'Dim sUTYPE As String = "" 'FType'保險種類。 (A表示勞保+就保，L表示勞保，V農保)
    'Dim sBDATE As String = "" '投保起日(YYYYMMDD西元年月日)
    'Dim sEDATE As String = "" '投保迄日(YYYYMMDD西元年月日)
    'Dim gCmdStr As String = "" '讀取傳入的外部參數
    'Dim strDetail As String = "" '取得未整理的XML資訊

    'Const Cst_Errmsg_1 As String = "查詢資料找不到" '不算錯誤
    'Const Cst_Errmsg_2 As String = "查詢資料格式不符"
    'Const Cst_Errmsg_3 As String = "不允許的查詢"
    'Const Cst_Errmsg_4 As String = "不明的錯誤"
    'Const Cst_Errmsg_5 As String = "程式內部錯誤"
    'Const Cst_Errmsg_6 As String = "錯誤碼:6"
    'Const Cst_Errmsg_7 As String = "身分證重號"
    'Const Cst_Errmsg_8 As String = "資料超過100筆"

    '確定
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String = ""
        Dim SB4ID As String = ""
        Dim SMDATE As String = ""
        Dim FMDATE As String = ""

        For Each eItem As DataGridItem In DataGrid1.Items
            Dim Radio1 As HtmlInputRadioButton = eItem.FindControl("Radio1")
            Dim Hid_sb4id As HiddenField = eItem.FindControl("Hid_sb4id")
            Dim Hid_SMDATE As HiddenField = eItem.FindControl("Hid_SMDATE")
            Dim Hid_FMDATE As HiddenField = eItem.FindControl("Hid_FMDATE")
            If Radio1.Checked AndAlso Hid_sb4id.Value <> "" Then
                SB4ID = Hid_sb4id.Value
                SMDATE = Hid_SMDATE.Value
                FMDATE = Hid_FMDATE.Value
                Exit For
            End If
        Next

        If SB4ID = "" Then
            Common.MessageBox(Me, "查無資料，無法回傳值")
            Exit Sub
        End If

        Dim drSB4 As DataRow = TIMS.Get_BLIGATEDATA4(SB4ID, Hid_idno.Value, objconn)
        If drSB4 Is Nothing Then
            Common.MessageBox(Me, "查無資料，無法回傳值")
            Exit Sub
        End If

        Dim Script As String = ""
        Select Case Hid_SPAGE.Value
            Case cst_SPAGE_SD01001
                'getRadioValue(document.form1.PriorWorkType1) 
                'var PWTvalue1 = getRadioValue(document.form1.PriorWorkType1);
                '//alert(PWTvalue1);
                'if (PWTvalue1 == '2') {
                '	var PriorWorkType1 = document.getElementById("PriorWorkType1");
                '	setValue(PriorWorkType1, "1");
                '	//setSPANValue(PriorWorkType1, "");
                '}
                Script = ""
                Script &= "<script>" & vbCrLf
                Script &= "var hidSB4ID = opener.document.getElementById('hidSB4ID');" & vbCrLf
                Script &= "var PriorWorkType1 = opener.document.getElementById('PriorWorkType1');" & vbCrLf
                Script &= "var PriorWorkOrg1 = opener.document.getElementById('PriorWorkOrg1');" & vbCrLf
                Script &= "var ActNo = opener.document.getElementById('ActNo');" & vbCrLf
                Script &= "var SOfficeYM1 = opener.document.getElementById('SOfficeYM1');" & vbCrLf
                Script &= "var FOfficeYM1 = opener.document.getElementById('FOfficeYM1');" & vbCrLf
                Script &= "var PWTvalue1 = getRadioValue(opener.document.form1.PriorWorkType1);" & vbCrLf
                Script &= "if(PWTvalue1 == '2'){" & vbCrLf
                Script &= " setValue(PriorWorkType1, '1');" & vbCrLf
                Script &= "}" & vbCrLf
                Script &= "hidSB4ID.value='" & SB4ID & "';" & vbCrLf
                Script &= "PriorWorkOrg1.value='" & drSB4("comname") & "';" & vbCrLf
                Script &= "ActNo.value='" & drSB4("actno") & "';" & vbCrLf
                Script &= "PriorWorkOrg1.disabled=true;" & vbCrLf
                Script &= "ActNo.disabled=true;" & vbCrLf
                'Script &= "PriorWorkOrg1.readOnly=true;" & vbCrLf
                'Script &= "ActNo.readOnly=true;" & vbCrLf
                Select Case Convert.ToString(drSB4("changemode"))
                    Case "4" '加保
                        Script &= "SOfficeYM1.value='" & SMDATE & "';" & vbCrLf
                        Script &= "FOfficeYM1.value='" & FMDATE & "';" & vbCrLf
                        'Script &= "SOfficeYM1.value='" & drSB4("mdate_ad") & "';" & vbCrLf
                        'Script &= "FOfficeYM1.value='';" & vbCrLf
                    Case "2" '退保
                        Script &= "SOfficeYM1.value='" & SMDATE & "';" & vbCrLf
                        Script &= "FOfficeYM1.value='" & FMDATE & "';" & vbCrLf
                        'Script &= "SOfficeYM1.value='';" & vbCrLf
                        'Script &= "FOfficeYM1.value='" & drSB4("mdate_ad") & "';" & vbCrLf
                End Select
                Script &= "window.top.opener = null;" & vbCrLf
                Script &= "window.close();" & vbCrLf
                Script &= "</script>" & vbCrLf
            Case cst_SPAGE_SD03002
                Script = ""
                Script &= "<script>" & vbCrLf
                Script &= "var hidSB4ID = opener.document.getElementById('hidSB4ID');" & vbCrLf
                Script &= "var PriorWorkType1 = opener.document.getElementById('PriorWorkType1');" & vbCrLf
                Script &= "var PriorWorkOrg1 = opener.document.getElementById('PriorWorkOrg1');" & vbCrLf
                Script &= "var ActNo2 = opener.document.getElementById('ActNo2');" & vbCrLf
                Script &= "var PriorWorkPay = opener.document.getElementById('PriorWorkPay');" & vbCrLf
                Script &= "var SOfficeYM1 = opener.document.getElementById('SOfficeYM1');" & vbCrLf
                Script &= "var FOfficeYM1 = opener.document.getElementById('FOfficeYM1');" & vbCrLf
                Script &= "var PWTvalue1 = getRadioValue(opener.document.form1.PriorWorkType1);" & vbCrLf
                Script &= "if(PWTvalue1 == '2'){" & vbCrLf
                Script &= " setValue(PriorWorkType1, '1');" & vbCrLf
                Script &= "}" & vbCrLf
                Script &= "hidSB4ID.value='" & SB4ID & "';" & vbCrLf
                Script &= "PriorWorkOrg1.value='" & drSB4("comname") & "';" & vbCrLf
                Script &= "ActNo2.value='" & drSB4("actno") & "';" & vbCrLf
                Script &= "PriorWorkPay.value='" & drSB4("SALARY") & "';" & vbCrLf
                Script &= "PriorWorkOrg1.disabled=true;" & vbCrLf
                Script &= "ActNo2.disabled=true;" & vbCrLf
                Script &= "PriorWorkPay.disabled=true;" & vbCrLf
                'Script &= "PriorWorkOrg1.readOnly=true;" & vbCrLf
                'Script &= "ActNo2.readOnly=true;" & vbCrLf
                'Script &= "PriorWorkPay.readOnly=true;" & vbCrLf
                Select Case Convert.ToString(drSB4("changemode"))
                    Case "4" '加保
                        Script &= "SOfficeYM1.value='" & SMDATE & "';" & vbCrLf
                        Script &= "FOfficeYM1.value='" & FMDATE & "';" & vbCrLf
                        'Script &= "SOfficeYM1.value='" & drSB4("mdate_ad") & "';" & vbCrLf
                        'Script &= "FOfficeYM1.value='';" & vbCrLf
                    Case "2" '退保
                        Script &= "SOfficeYM1.value='" & SMDATE & "';" & vbCrLf
                        Script &= "FOfficeYM1.value='" & FMDATE & "';" & vbCrLf
                        'Script &= "SOfficeYM1.value='';" & vbCrLf
                        'Script &= "FOfficeYM1.value='" & drSB4("mdate_ad") & "';" & vbCrLf
                End Select
                Script &= "window.top.opener = null;" & vbCrLf
                Script &= "window.close();" & vbCrLf
                Script &= "</script>"
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Exit Sub
                'Script = ""
                'Script &= "<script>" & vbCrLf
                'Script &= "opener.document.getElementById('PriorWorkOrg1').value='" & dr("comname") & "';" & vbCrLf
                'Script &= "opener.document.getElementById('ActNo').value='" & dr("actno") & "';" & vbCrLf
                'Select Case Convert.ToString(dr("changemode"))
                '    Case "4" '加保
                '        Script &= "opener.document.getElementById('SOfficeYM1').value='" & dr("mdate_ad") & "';" & vbCrLf
                '        Script &= "opener.document.getElementById('FOfficeYM1').value='';" & vbCrLf
                '    Case "2" '退保
                '        Script &= "opener.document.getElementById('SOfficeYM1').value='';" & vbCrLf
                '        Script &= "opener.document.getElementById('FOfficeYM1').value='" & dr("mdate_ad") & "';" & vbCrLf
                'End Select
                'Script &= "window.top.opener = null;" & vbCrLf
                'Script &= "window.close();" & vbCrLf
                'Script &= "</script>"
        End Select

        Page.RegisterStartupScript(TIMS.xBlockName(), Script)

    End Sub

End Class

