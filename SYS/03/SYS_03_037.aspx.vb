Imports System.Threading.Tasks

Partial Class SYS_03_037
    Inherits AuthBasePage

    Const cst_TipMsg1 As String = "已排除 預算別再出發!!"
    Const cst_inline1 As String = ""
    Const cst_none1 As String = "none"

    Const cst_StatusMsg1_Y As String = "審核通過"
    Const cst_StatusMsg1_N As String = "審核不通過"
    Const cst_StatusMsg1_R As String = "退件修正"
    Const cst_StatusMsg1_S As String = "審核中"
    Const cst_StatusMsg1_NOINFO As String = "無資訊"
    Const cst_AppliedStatusM_NOINFO As String = "NOINFO"

    Const cst_StatusMsg2_1 As String = "已撥款"
    Const cst_StatusMsg2_Y As String = "待撥款" '"撥款中"
    Const cst_StatusMsg2_N As String = "不撥款"
    Const cst_StatusMsg2_R As String = "未撥款"
    Const cst_StatusMsg2_X As String = "不予補助"
    Const cst_StatusMsg2_NOINFO As String = "無資訊"

    Dim dtIdentity As DataTable = Nothing
    Dim dtTrade As DataTable = Nothing
    Dim dtZip As DataTable = Nothing

    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objConn)

        If sm.UserInfo.LID <> 0 Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg5s)
            ShowScreen1(0)
            Return
        End If

        If Not IsPostBack Then
            CCREATE1()
        End If
    End Sub

    Sub CCREATE1()
        lab_Msg.Text = ""
        ShowScreen1(0)
        Dim V_INQUIRY As String = Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me)))
        If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objConn, V_INQUIRY)

        Dim MRqID As String = TIMS.Get_MRqID(Me)
        TIMS.Get_TitleLab(objConn, MRqID, TitleLab1, TitleLab2)
    End Sub

    '查詢
    Private Sub btn_Query_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Query.Click
        If sm.UserInfo.LID <> 0 Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg5s)
            ShowScreen1(0)
            Return
        End If

        txt_IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(txt_IDNO.Text))
        txt_NAME.Text = TIMS.ClearSQM(txt_NAME.Text)
        If String.IsNullOrEmpty(txt_IDNO.Text) AndAlso String.IsNullOrEmpty(txt_NAME.Text) Then
            Common.MessageBox(Me, "請輸入查詢值! (身分證號 或 姓名)")
            Return
        End If

        Dim myValue1 As String = ""
        TIMS.SetMyValue(myValue1, "IDNO", txt_IDNO.Text)
        TIMS.SetMyValue(myValue1, "NAME", txt_NAME.Text)
        Hid_sch1.Value = TIMS.EncryptAes(myValue1)

        Call sSearch1(0)
    End Sub

    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        txt_IDNO.Text = TIMS.ClearSQM(txt_IDNO.Text)
        txt_NAME.Text = TIMS.ClearSQM(txt_NAME.Text)
        If txt_IDNO.Text <> "" Then RstMemo &= String.Concat("&身分證號碼=", txt_IDNO.Text)
        If txt_NAME.Text <> "" Then RstMemo &= String.Concat("&Name=", txt_NAME.Text)
        Return RstMemo
    End Function


    Private Sub sSearch1(ByVal tmpPage As Integer)
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If
        'Call TIMS.sUtl_TxtPageSize(Me, TxtPageSize, DataGrid2)

        ShowScreen1(0)
        If Hid_sch1.Value = "" Then Return

        Dim MySchValue1 As String = TIMS.DecryptAes(Hid_sch1.Value)
        Dim vIDNO As String = TIMS.GetMyValue(MySchValue1, "IDNO")
        Dim vNAME As String = TIMS.GetMyValue(MySchValue1, "NAME")
        If vIDNO = "" AndAlso vNAME = "" Then Return

        Dim sParms As New Hashtable
        If vIDNO <> "" Then sParms.Add("IDNO", vIDNO)
        If vNAME <> "" Then sParms.Add("NAME", vNAME)

        'AND cs.IDNO='G220540773' 'SINFO','STEMP','STEMP2'
        Dim sSql As String = ""
        sSql = "" & vbCrLf
        sSql &= " WITH WS1 AS (SELECT IDNO,NAME,BIRTHDAY,SEX,'SINFO' MSTYLE FROM STUD_STUDENTINFO WITH(NOLOCK) WHERE 1=1" & vbCrLf
        sSql &= String.Concat(If(vIDNO <> "", " AND IDNO=@IDNO", ""), If(vNAME <> "", " AND NAME=@NAME", "")) & vbCrLf
        sSql &= " UNION SELECT IDNO,NAME,BIRTHDAY,SEX,'STEMP' MSTYLE FROM STUD_ENTERTEMP WITH(NOLOCK) WHERE 1=1" & vbCrLf
        sSql &= String.Concat(If(vIDNO <> "", " AND IDNO=@IDNO", ""), If(vNAME <> "", " AND NAME=@NAME", "")) & vbCrLf
        sSql &= " UNION SELECT IDNO,NAME,BIRTHDAY,SEX,'STEMP2' MSTYLE FROM STUD_ENTERTEMP2 WITH(NOLOCK) WHERE 1=1" & vbCrLf
        sSql &= String.Concat(If(vIDNO <> "", " AND IDNO=@IDNO", ""), If(vNAME <> "", " AND NAME=@NAME", ""), ")") & vbCrLf

        sSql &= " ,WS2 AS (SELECT IDNO, MAX(MSTYLE) MSTYLE FROM WS1 GROUP BY IDNO)" & vbCrLf

        sSql &= " SELECT ss.IDNO,ss.NAME STNAME" & vbCrLf
        sSql &= " ,FORMAT(ss.BIRTHDAY,'yyyy/MM/dd') BIRTHDAY,ss.SEX" & vbCrLf
        sSql &= " ,case ss.SEX when 'M' then '男' when 'F' then '女' end SEX2" & vbCrLf
        sSql &= " ,format(dbo.FN_GET_MODIFYDATE2(ss.MSTYLE,ss.IDNO),'yyyy/MM/dd HH:mm:ss') MODIFYDATE" & vbCrLf
        sSql &= " FROM WS2 s2" & vbCrLf
        sSql &= " JOIN WS1 ss ON ss.IDNO=s2.IDNO AND ss.MSTYLE=s2.MSTYLE" & vbCrLf
        'sSql &= " WHERE 1=1" & vbCrLf 'AND cs.IDNO='G220540773'
        'If vIDNO <> "" Then sSql &= " AND ss.IDNO=@IDNO" & vbCrLf 'If vNAME <> "" Then sSql &= " AND ss.NAME=@NAME" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objConn, sParms)

        Dim sMemo As String = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "STNAME,IDNO,BIRTHDAY,SEX2,MODIFYDATE")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip2, "", sMemo, objConn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        DataGrid2.Visible = False
        lab_Msg.Text = TIMS.cst_NODATAMsg1

        If TIMS.dtNODATA(dt) Then Return

        DataGrid2.Visible = True
        lab_Msg.Text = ""
        DataGrid2.DataSource = dt
        DataGrid2.CurrentPageIndex = tmpPage
        DataGrid2.DataBind()
    End Sub

    Private Sub DataGrid2_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        Dim sCmdName As String = If(e IsNot Nothing, e.CommandName, "")
        Dim sCmdArg As String = If(e IsNot Nothing, e.CommandArgument, "")
        If sCmdName = "" Then Return
        If sCmdArg = "" Then Return

        Select Case e.CommandName
            Case "BTNVIEW1"
                Hid_IDNO.Value = TIMS.GetMyValue(sCmdArg, "IDNO")
                If Hid_IDNO.Value = "" Then Return
                SHOW_DATA1(Hid_IDNO.Value)

                Show_DataGrid11(Hid_IDNO.Value)
                Show_DataGrid11b(Hid_IDNO.Value)

                Show_DataGrid12(Hid_IDNO.Value)
                Show_DataGrid12b(Hid_IDNO.Value)

                Show_DataGrid13(Hid_IDNO.Value)

                Show_DataGrid14(Hid_IDNO.Value)

                ShowScreen1(1)
                'Page.RegisterStartupScript("ChangeMode1", "<script>ChangeMode(1);</script>")
                TIMS.RegisterStartupScript(Me, "ChangeMode1", "<script>ChangeMode(1);</script>")
        End Select
    End Sub

    ''' <summary>清理資料</summary>
    Private Sub CLEAR_DATA1()
        LName.Text = "" 'Convert.ToString(dr1("STNAME"))
        LIDNO.Text = "" 'Convert.ToString(dr1("IDNO"))
        LBIRTH.Text = "" 'Convert.ToString(dr1("BIRTHDAY"))
        LSEX2.Text = "" 'Convert.ToString(dr1("SEX2"))
        Hid_MSTYLE.Value = "" ' Convert.ToString(dr1("MSTYLE"))

        PassPortNO.Text = TIMS.cst_NODATAMsg12 'If(Convert.ToString(dr1("PASSPORTNO")) = "1", "本國", "外國")
        IDNO.Text = TIMS.cst_NODATAMsg12 'Convert.ToString(dr1("IDNO"))
        Sex.Text = TIMS.cst_NODATAMsg12 'If(Convert.ToString(dr1("SEX")) = "M", "男", If(Convert.ToString(dr1("SEX")) = "F", "女", Convert.ToString(dr1("SEX"))))
        DegreeID.Text = TIMS.cst_NODATAMsg12 'TIMS.Get_DegreeValue(Convert.ToString(dr1("DEGREEID")), objConn)

        MaritalStatus.Text = TIMS.cst_NODATAMsg12 'If(Convert.ToString(dr1("MARITALSTATUS")) = "1", "已婚", If(Convert.ToString(dr1("MaritalStatus")) = "2", "未婚", TIMS.cst_NODATAMsg12))
        GradID.Text = TIMS.cst_NODATAMsg12 'TIMS.Get_GRADNAME(Convert.ToString(dr1("GRADID")), objConn)
        School.Text = TIMS.cst_NODATAMsg12 'Convert.ToString(dr1("SCHOOL"))
        Department.Text = TIMS.cst_NODATAMsg12 'Convert.ToString(dr1("DEPARTMENT"))
        MilitaryID.Text = TIMS.cst_NODATAMsg12 'TIMS.Get_MILITARYNAME(Convert.ToString(dr1("MILITARYID")), objConn)

        Address.Text = TIMS.cst_NODATAMsg12 'TIMS.getZipName6(s_ZipCODE, sAddress, "", dtZip)
        LabHouseholdAddress.Text = TIMS.cst_NODATAMsg12 'Convert.ToString(dr1("STNAME"))

        Phone1.Text = TIMS.cst_NODATAMsg12 ' Convert.ToString(dr1("PHONE1"))
        Phone2.Text = TIMS.cst_NODATAMsg12 ' Convert.ToString(dr1("PHONE2"))
        Email.Text = TIMS.cst_NODATAMsg12 ' Convert.ToString(dr1("EMAIL"))
        CellPhone.Text = TIMS.cst_NODATAMsg12 ' Convert.ToString(dr1("CELLPHONE"))
    End Sub

    ''' <summary>個人基本資料</summary>
    ''' <param name="vIDNO"></param>
    Private Sub SHOW_DATA1(vIDNO As String)
        CLEAR_DATA1()
        Dim sParms As New Hashtable
        sParms.Add("IDNO", vIDNO)
        Dim sSql As String = ""

        sSql = "" & vbCrLf
        sSql &= " WITH WS1 AS (SELECT IDNO,NAME,BIRTHDAY,SEX,'SINFO' MSTYLE FROM STUD_STUDENTINFO WITH(NOLOCK) WHERE 1=1" & vbCrLf
        sSql &= String.Concat(If(vIDNO <> "", " AND IDNO=@IDNO", ""), "") & vbCrLf
        sSql &= " UNION SELECT IDNO,NAME,BIRTHDAY,SEX,'STEMP' MSTYLE FROM STUD_ENTERTEMP WITH(NOLOCK) WHERE 1=1" & vbCrLf
        sSql &= String.Concat(If(vIDNO <> "", " AND IDNO=@IDNO", ""), "") & vbCrLf
        sSql &= " UNION SELECT IDNO,NAME,BIRTHDAY,SEX,'STEMP2' MSTYLE FROM STUD_ENTERTEMP2 WITH(NOLOCK) WHERE 1=1" & vbCrLf
        sSql &= String.Concat(If(vIDNO <> "", " AND IDNO=@IDNO", ""), ")") & vbCrLf

        sSql &= " SELECT ss.IDNO,ss.NAME SNAME" & vbCrLf
        sSql &= " ,dbo.FN_GET_MASK1(ss.IDNO) IDNO_MK" & vbCrLf
        sSql &= " ,dbo.FN_GET_MASK2(ss.BIRTHDAY) BIRTHDAY_MK" & vbCrLf
        sSql &= " ,dbo.FN_GET_MASK3(ss.NAME) NAME_MK" & vbCrLf
        sSql &= " ,FORMAT(ss.BIRTHDAY,'yyyy/MM/dd') BIRTHDAY,ss.SEX" & vbCrLf
        sSql &= " ,case ss.SEX when 'M' then '男' when 'F' then '女' end SEX2" & vbCrLf
        sSql &= " ,ss.MSTYLE" & vbCrLf
        'sSql &= " ,format(dbo.FN_GET_MODIFYDATE2(ss.MSTYLE,ss.IDNO),'yyyy/MM/dd HH:mm:ss') MODIFYDATE" & vbCrLf
        'sSql &= " ,format(ss.MODIFYDATE,'yyyy/MM/dd HH:mm:ss') MODIFYDATE" & vbCrLf
        sSql &= " FROM WS1 ss" & vbCrLf
        'sSql &= " WHERE 1=1" & vbCrLf ' AND cs.IDNO='G220540773'
        'sSql &= " AND ss.IDNO=@IDNO" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objConn, sParms)

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return
        Dim dr1 As DataRow = dt.Rows(0)
        LName.Text = Convert.ToString(dr1("SNAME"))
        LIDNO.Text = Convert.ToString(dr1("IDNO_MK"))
        LBIRTH.Text = Convert.ToString(dr1("BIRTHDAY_MK"))
        LSEX2.Text = Convert.ToString(dr1("SEX2"))
        Hid_MSTYLE.Value = Convert.ToString(dr1("MSTYLE"))
    End Sub

    ''' <summary> 個人基本資料-戶籍地址-</summary>
    ''' <param name="vIDNO"></param>
    Private Sub Show_DataGrid11b(vIDNO As String)
        Dim tPrams As New Hashtable
        tPrams.Add("IDNO", vIDNO)
        Dim sSql As String = ""
        sSql = "" & vbCrLf
        sSql &= " WITH WC1 AS (SELECT MAX(c.ESERNUM) ESERNUM FROM STUD_ENTERTEMP2 a" & vbCrLf
        sSql &= "  JOIN STUD_ENTERTYPE2 b on b.ESETID=a.ESETID" & vbCrLf
        sSql &= "  JOIN STUD_ENTERTRAIN2 c on c.ESERNUM=b.ESERNUM WHERE a.IDNO=@IDNO)" & vbCrLf 'A123456789'

        sSql &= " SELECT a.SEID,a.ESERNUM,a.ZIPCODE2,a.HOUSEHOLDADDRESS,a.ZIPCODE2_6W" & vbCrLf
        sSql &= " FROM WC1 c JOIN STUD_ENTERTRAIN2 a ON a.ESERNUM=c.ESERNUM" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objConn, tPrams)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return
        Dim dr1 As DataRow = dt.Rows(0)

        If dtZip Is Nothing Then dtZip = TIMS.Get_VZipName(objConn)

        Dim sAddress As String = Convert.ToString(dr1("HOUSEHOLDADDRESS"))
        'If Session(TIMS.gcst_rblWorkMode) = TIMS.cst_wmdip1 Then sAddress = TIMS.strMask(sAddress, 3)
        Dim s_ZipCode2 As String = If(Convert.ToString(dr1("ZIPCODE2_6W")) <> "", Convert.ToString(dr1("ZIPCODE2_6W")), Convert.ToString(dr1("ZipCode2")))
        Dim s_HouseholdAddress As String = TIMS.getZipName6(s_ZipCode2, sAddress, "", dtZip)
        'Label15.Text = s_HouseholdAddress
        LabHouseholdAddress.Text = s_HouseholdAddress ' Convert.ToString(dr1("STNAME"))
    End Sub

    ''' <summary>個人基本資料</summary>
    ''' <param name="vIDNO"></param>
    Private Sub Show_DataGrid11(vIDNO As String)
        'Dim dt As DataTable = TIMS.dtNothing
        Dim tPrams As New Hashtable
        tPrams.Add("IDNO", vIDNO)
        Dim sSql As String = ""
        sSql = "" & vbCrLf
        sSql &= " WITH WC1 AS (SELECT MAX(ESETID) ESETID FROM STUD_ENTERTEMP2 WHERE IDNO=@IDNO)" & vbCrLf 'A123456789'

        sSql &= " SELECT a.ESETID,a.SETID ,a.IDNO ,a.NAME ANAME,a.SEX ,a.BIRTHDAY ,a.PASSPORTNO ,a.MARITALSTATUS ,a.DEGREEID" & vbCrLf
        sSql &= " ,a.GRADID ,a.SCHOOL ,a.DEPARTMENT ,a.MILITARYID ,a.ZIPCODE ,a.ADDRESS ,a.PHONE1 ,a.PHONE2" & vbCrLf
        sSql &= " ,a.CELLPHONE ,a.EMAIL ,a.ISAGREE ,a.ZIPCODE6W" & vbCrLf
        sSql &= " ,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK" & vbCrLf
        'sSql &= " ,dbo.FN_GET_MASK2(a.BIRTHDAY) BIRTHDAY_MK" & vbCrLf
        sSql &= " ,dbo.FN_GET_MASK3(a.NAME) NAME_MK" & vbCrLf
        sSql &= " FROM WC1 c JOIN STUD_ENTERTEMP2 a on a.ESETID=c.ESETID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objConn, tPrams)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return
        Dim dr1 As DataRow = dt.Rows(0)

        If dtZip Is Nothing Then dtZip = TIMS.Get_VZipName(objConn)

        PassPortNO.Text = If(Convert.ToString(dr1("PASSPORTNO")) = "1", "本國", "外國")
        IDNO.Text = Convert.ToString(dr1("IDNO_MK"))
        Sex.Text = If(Convert.ToString(dr1("SEX")) = "M", "男", If(Convert.ToString(dr1("SEX")) = "F", "女", Convert.ToString(dr1("SEX"))))
        DegreeID.Text = TIMS.Get_DegreeValue(Convert.ToString(dr1("DEGREEID")), objConn)

        MaritalStatus.Text = If(Convert.ToString(dr1("MARITALSTATUS")) = "1", "已婚", If(Convert.ToString(dr1("MaritalStatus")) = "2", "未婚", TIMS.cst_NODATAMsg12))
        GradID.Text = TIMS.Get_GRADNAME(Convert.ToString(dr1("GRADID")), objConn)
        School.Text = Convert.ToString(dr1("SCHOOL"))
        Department.Text = Convert.ToString(dr1("DEPARTMENT"))
        MilitaryID.Text = TIMS.Get_MILITARYNAME(Convert.ToString(dr1("MILITARYID")), objConn)

        Dim sAddress As String = Convert.ToString(dr1("Address"))
        'If Session(TIMS.gcst_rblWorkMode) = TIMS.cst_wmdip1 Then sAddress = TIMS.strMask(sAddress, 3)
        Dim s_ZipCODE As String = If(Convert.ToString(dr1("ZipCODE6W")) <> "", Convert.ToString(dr1("ZipCODE6W")), Convert.ToString(dr1("ZipCODE")))
        Address.Text = TIMS.getZipName6(s_ZipCODE, sAddress, "", dtZip)
        'Address.Text = Convert.ToString(dr1("ADDRESS"))
        'LabHouseholdAddress.Text = Convert.ToString(dr1("STNAME"))

        If (Convert.ToString(dr1("PHONE1")) <> "") Then Phone1.Text = Convert.ToString(dr1("PHONE1"))
        If (Convert.ToString(dr1("PHONE2")) <> "") Then Phone2.Text = Convert.ToString(dr1("PHONE2"))
        If (Convert.ToString(dr1("EMAIL")) <> "") Then Email.Text = Convert.ToString(dr1("EMAIL"))
        If (Convert.ToString(dr1("CELLPHONE")) <> "") Then CellPhone.Text = Convert.ToString(dr1("CELLPHONE"))

        'lab_Msg11.Visible = True
        'Datagrid11.Visible = False
        'If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        'lab_Msg11.Visible = False
        'Datagrid11.Visible = True
        'Datagrid11.DataSource = dt
        'Datagrid11.DataBind()
    End Sub

    ''' <summary>曾報名課程-已報名課程</summary>
    ''' <param name="vIDNO"></param>
    Private Sub Show_DataGrid12(vIDNO As String)
        Dim nPrams As New Hashtable
        nPrams.Add("IDNO", vIDNO)
        Dim sSql As String = ""
        sSql = "" & vbCrLf
        sSql &= " SELECT cc.YEARS,cc.PLANNAME,cc.DISTNAME" & vbCrLf
        sSql &= " ,cc.ORGNAME,ISNULL(cc.OCID,b.OCID1) OCID" & vbCrLf
        sSql &= " ,cc.CLASSCNAME2" & vbCrLf
        sSql &= " ,concat(format(cc.STDATE,'yyyy/MM/dd'),'-',format(cc.FTDATE,'yyyy/MM/dd')) SFTDATE" & vbCrLf
        sSql &= " ,format(b.RELENTERDATE,'yyyy/MM/dd HH:mm:ss') RELENTERDATE" & vbCrLf
        sSql &= " ,b.ESERNUM,b.ESETID,b.SIGNUPSTATUS" & vbCrLf
        sSql &= " ,case b.SIGNUPSTATUS when 0 then '尚未審核' when 2 then '審核失敗' else case when b.SIGNUPSTATUS IN (1,3,4,5) then '審核成功' else concat('其他狀況.',b.SIGNUPSTATUS) end end SIGNUPSTATUS_N"
        sSql &= " FROM STUD_ENTERTEMP2 a" & vbCrLf
        sSql &= " JOIN STUD_ENTERTYPE2 b on a.ESETID=b.ESETID" & vbCrLf
        sSql &= " JOIN VIEW2 cc on cc.OCID=b.OCID1" & vbCrLf
        sSql &= " WHERE a.IDNO=@IDNO" & vbCrLf 'A222929936'" & vbCrLf
        sSql &= " ORDER BY cc.YEARS,cc.STDATE" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objConn, nPrams)

        lab_Msg12.Visible = True
        Datagrid12.Visible = False
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        lab_Msg12.Visible = False
        Datagrid12.Visible = True
        Datagrid12.DataSource = dt
        Datagrid12.DataBind()
    End Sub

    ''' <summary>曾報名課程-取消報名課程</summary>
    ''' <param name="vIDNO"></param>
    Private Sub Show_DataGrid12b(vIDNO As String)
        Dim nPrams As New Hashtable
        nPrams.Add("IDNO", vIDNO)
        Dim sSql As String = ""
        sSql = "" & vbCrLf
        sSql &= " SELECT cc.YEARS,cc.PLANNAME,cc.DISTNAME" & vbCrLf
        sSql &= " ,cc.ORGNAME,ISNULL(cc.OCID,b.OCID1) OCID" & vbCrLf
        sSql &= " ,cc.CLASSCNAME2" & vbCrLf
        sSql &= " ,concat(format(cc.STDATE,'yyyy/MM/dd'),'-',format(cc.FTDATE,'yyyy/MM/dd')) SFTDATE" & vbCrLf
        sSql &= " ,format(b.RELENTERDATE,'yyyy/MM/dd HH:mm:ss') RELENTERDATE" & vbCrLf
        sSql &= " ,b.ESERNUM,b.ESETID,b.SIGNUPSTATUS" & vbCrLf
        sSql &= " ,format(b.MODIFYDATE,'yyyy/MM/dd HH:mm:ss') CANCELTIME" & vbCrLf
        sSql &= " FROM STUD_ENTERTEMP2 a" & vbCrLf
        sSql &= " JOIN STUD_ENTERTYPE2DELDATA b on a.ESETID=b.ESETID and a.IDNO=b.MODIFYACCT" & vbCrLf
        sSql &= " JOIN VIEW2 cc on cc.OCID=b.OCID1" & vbCrLf
        sSql &= " WHERE a.IDNO=@IDNO" & vbCrLf 'A222929936'" & vbCrLf B222911659,L220226708,D122001649,G221221133,B222911659,
        sSql &= " ORDER BY cc.YEARS,cc.STDATE" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objConn, nPrams)

        lab_Msg12b.Visible = True
        Datagrid12b.Visible = False
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        lab_Msg12b.Visible = False
        Datagrid12b.Visible = True
        Datagrid12b.DataSource = dt
        Datagrid12b.DataBind()
    End Sub

    ''' <summary>參訓學員歷史</summary>
    ''' <param name="vIDNO"></param>
    Private Sub Show_DataGrid13(vIDNO As String)
        Dim dt As DataTable = TIMS.SchHistoryStudInfo(objConn, vIDNO)

        lab_Msg13.Visible = True
        Datagrid13.Visible = False
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return
        dt.DefaultView.Sort = "Years,TRound"

        lab_Msg13.Visible = False
        Datagrid13.Visible = True
        Datagrid13.DataSource = dt.DefaultView.Table
        Datagrid13.DataBind()
    End Sub

    ''' <summary>補助費用歷史 </summary>
    ''' <param name="vIDNO"></param>
    Private Sub Show_DataGrid14(vIDNO As String)
        'D120272346,C220037580,C220577258,Q220976976,C220534459,C220252265,C200851379,F123303297,C120572724,C220477897,
        Dim oParms As New Hashtable From {{"IDNO", vIDNO}}
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT ip.YEARS,ip.DISTNAME,rr.OrgName" & vbCrLf '機構名稱
        sql &= " ,c.OCID,dbo.FN_GET_CLASSCNAME(c.CLASSCNAME,c.CYCLTYPE) CLASSCNAME" & vbCrLf '班名
        sql &= " ,format(c.STDate,'yyyy/MM/dd') STDate" & vbCrLf '開訓日
        sql &= " ,format(c.FTDate,'yyyy/MM/dd') FTDate" & vbCrLf '結訓日
        sql &= " ,e.SumOfMoney " & vbCrLf
        sql &= " ,e.AppliedStatus " & vbCrLf
        sql &= " ,e.AppliedStatusM " & vbCrLf
        sql &= " ,e.BUDID " & vbCrLf
        sql &= " ,bb.BUDNAME " & vbCrLf '預算別
        sql &= " ,b.SOCID " & vbCrLf 'SOCID
        sql &= " ,b.StudStatus SDSTATUS " & vbCrLf 'StudStatus->SdStatus
        sql &= " ,dbo.DECODE12(b.StudStatus,1,'在訓',2,'離訓',3,'退訓',4,'續訓',5,'結訓','在訓') StudStatus " & vbCrLf 'StudStatus
        sql &= " FROM STUD_STUDENTINFO a " & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS b ON a.SID=b.SID " & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO c ON b.OCID=c.OCID " & vbCrLf
        sql &= " JOIN PLAN_PLANINFO d ON c.PlanID=d.PlanID AND c.ComIDNO=d.ComIDNO AND c.SeqNO=d.SeqNO " & vbCrLf
        'Cst_TPlanID28_1 '要含充電起飛的包班資料
        sql &= String.Concat(" AND d.TPlanID IN (", TIMS.Cst_TPlanID28_1a, ")") & vbCrLf

        sql &= " JOIN VIEW_RIDNAME rr ON rr.RID=c.RID " & vbCrLf
        sql &= " JOIN VIEW_PLAN ip ON ip.planid=c.planid "
        sql &= " LEFT JOIN dbo.STUD_QUESTIONFIN qq ON qq.socid=b.socid " & vbCrLf
        '學員經費 已撥款狀態 e.AppliedStatus=1
        sql &= " LEFT JOIN dbo.STUD_SUBSIDYCOST e ON b.SOCID=e.SOCID " & vbCrLf 'AND e.AppliedStatus=1
        sql &= " LEFT JOIN dbo.VIEW_BUDGET bb ON bb.budid=e.budid " & vbCrLf
        sql &= " WHERE d.AppliedResult='Y' AND c.NotOpen='N' " & vbCrLf
        sql &= String.Concat(" AND ip.TPlanID IN (", TIMS.Cst_TPlanID28_1a, ")") & vbCrLf
        sql &= " AND a.IDNO =@IDNO" & vbCrLf
        sql &= " ORDER BY d.PlanYear ,c.STDate ASC " & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objConn, oParms)

        Dim flag_need_sort_1 As Boolean = False '須要重新排序嗎 false:不用
        Dim dt2 As DataTable = TIMS.GetTrainingList2c(objConn, vIDNO)
        Dim ss3 As String = ""
        Dim ff3 As String = ""

        ff3 = "TPlanID IN (" & TIMS.Cst_TPlanID28_1b1b & ")"
        If dt2.Select(ff3).Length > 0 Then
            For Each dr2 As DataRow In dt2.Rows
                Dim dr1 As DataRow = dt.NewRow()
                dr1("YEARS") = dr2("YEARS")
                dr1("DISTNAME") = dr2("DISTNAME")
                dr1("OrgName") = dr2("OrgName")
                dr1("ClassCName") = String.Concat(dr2("ClassCName"), "-", dr2("PLANNAME"))
                dr1("STDate") = TIMS.Cdate3(dr2("STDate"))
                dr1("FTDate") = TIMS.Cdate3(dr2("FTDate"))
                dr1("SumOfMoney") = dr2("SumOfMoney")
                dr1("AppliedStatus") = dr2("AppliedStatus")
                '強制改為審核通過 
                'dr1("AppliedStatusM") = dr2("AppliedStatusM") 'TIMS.cst_YES 'dr2("AppliedStatusM")
                dr1("AppliedStatusM") = dr2("AppliedStatusM") 'TIMS.cst_YES 'dr2("AppliedStatusM")
                dr1("BUDID") = dr2("BUDID")
                dr1("BUDNAME") = TIMS.GET_BudgetName(Convert.ToString(dr2("BUDID")), objConn)
                dr1("SOCID") = Convert.DBNull
                dr1("SDSTATUS") = dr2("STUDSTATUS")
                dr1("StudStatus") = dr2("TFLAG")
                dt.Rows.Add(dr1)
            Next
            flag_need_sort_1 = True '須要重新排序嗎 true:必須
            dt.AcceptChanges()
        End If

        If flag_need_sort_1 Then
            ff3 = ""
            ss3 = "STDate"
            dt = TIMS.CopyDt(dt, ff3, ss3)
        End If

        '查無資料:TRUE
        Dim fgNODATA As Boolean = (dt Is Nothing OrElse dt.Rows.Count = 0)

        '(有資料才做)
        If Not fgNODATA Then
            '可用補助額'產投 政府補助經費 --產業人才投資方案(三年補助)
            RemainSub.Text = TIMS.Get_3Y_SupplyMoney()

            LabTotal.Text = RemainSub.Text
            LabTotal.ToolTip = TIMS.gTip_LabTotalSupplyMoney
            LabSumOfMoney.Text = 0

            Dim STDate As String = TIMS.Cdate3(Now)
            If dt.Rows.Count > 0 Then STDate = CDate(dt.Rows(dt.Rows.Count - 1)("STDate")).ToString("yyyy/MM/dd")
            '每3年所使用的補助金 目前已使用多少政府補助。
            '含職前webservice
            LabSumOfMoney.Text += TIMS.Get_SubsidyCost(vIDNO, STDate, "", "Y", objConn)
            TIMS.Tooltip(Me.LabSumOfMoney, cst_TipMsg1, True)
            RemainSub.Text = Int(RemainSub.Text) - CInt(Me.LabSumOfMoney.Text)

            Dim sDate As String = String.Empty
            Dim eDate As String = String.Empty
            'Dim aIDNO As String = e.CommandArgument
            Call TIMS.Get_SubSidyCostDay(vIDNO, STDate, sDate, eDate, objConn)
            TIMS.Tooltip(RemainSub, "計算開始日為：" & sDate & "~" & eDate, True)
            LabCostDay.Text = String.Concat("補助金補助期間：", sDate, "~", eDate)

            RemainSub.ForeColor = Color.Black
            If Int(RemainSub.Text) < 0 Then RemainSub.ForeColor = Color.Red
        End If

        '查無資料:TRUE
        'Dim fgNODATA As Boolean = (dt Is Nothing OrElse dt.Rows.Count = 0)
        LabTotal.Visible = If(fgNODATA, False, True)
        LabSumOfMoney.Visible = If(fgNODATA, False, True)
        RemainSub.Visible = If(fgNODATA, False, True)
        LabCostDay.Visible = If(fgNODATA, False, True)

        lab_Msg14.Visible = True
        Datagrid14.Visible = False

        If fgNODATA Then Return
        lab_Msg14.Visible = False
        Datagrid14.Visible = True

        Datagrid14.DataSource = dt
        Datagrid14.DataBind()
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        If e Is Nothing Then Return
        If e.Item.ItemType <> ListItemType.AlternatingItem AndAlso e.Item.ItemType <> ListItemType.Item Then Return

        Dim dr_Data As DataRowView = e.Item.DataItem
        Dim labSNo As Label = e.Item.FindControl("lab_SNo")
        Dim labIDNO As Label = e.Item.FindControl("lab_IDNO")
        Dim labName As Label = e.Item.FindControl("lab_Name")
        Dim labBirthday As Label = e.Item.FindControl("lab_Birthday")
        Dim lab_SexN As Label = e.Item.FindControl("lab_SexN")
        Dim lab_LastDate As Label = e.Item.FindControl("lab_LastDate")

        Dim BTNVIEW1 As LinkButton = e.Item.FindControl("BTNVIEW1") 'BTNVIEW1

        'labSNo.Text = Convert.ToString(DataGrid2.CurrentPageIndex * DataGrid2.PageSize + e.Item.ItemIndex + 1)
        labSNo.Text = TIMS.Get_DGSeqNo(sender, e)
        labIDNO.Text = Convert.ToString(dr_Data("IDNO"))
        labName.Text = Convert.ToString(dr_Data("STNAME"))
        labBirthday.Text = Convert.ToString(dr_Data("BIRTHDAY"))
        lab_SexN.Text = Convert.ToString(dr_Data("SEX2"))
        lab_LastDate.Text = Convert.ToString(dr_Data("MODIFYDATE"))

        'If BTNVIEW1 IsNot Nothing Then BTNVIEW1.Visible = False
        Dim sCmdArg As String = ""
        TIMS.SetMyValue(sCmdArg, "IDNO", Convert.ToString(dr_Data("IDNO")))
        BTNVIEW1.CommandArgument = sCmdArg
    End Sub

    Private Sub DataGrid2_PageIndexChanged(source As Object, e As DataGridPageChangedEventArgs) Handles DataGrid2.PageIndexChanged
        sSearch1(e.NewPageIndex)
    End Sub

    ''' <summary>顯示狀況</summary>
    ''' <param name="iPage"></param>
    Sub ShowScreen1(ByVal iPage As Integer)
        divSch1.Visible = False '(查詢)
        divShowData1.Visible = False '(資料顯示)
        Select Case iPage
            Case 0 '(查詢)
                divSch1.Visible = True
            Case 1 '(資料顯示)
                divShowData1.Visible = True
                '$("#MenuTable_td_1").addClass("active");
                MenuTable_td_1.Attributes("class") = "active"
                tb_VIEW1.Style("display") = cst_inline1
                tb_VIEW2.Style("display") = cst_none1
                tb_VIEW3.Style("display") = cst_none1
                tb_VIEW4.Style("display") = cst_none1
        End Select
    End Sub

    ''' <summary>回上頁</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnBACK1_Click(sender As Object, e As EventArgs) Handles btnBACK1.Click
        ShowScreen1(0)
    End Sub

    'Private Sub Datagrid11_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles Datagrid11.ItemDataBound
    '    If e Is Nothing Then Return
    '    Dim labSNo As Label = e.Item.FindControl("lab_SNo11")
    '    If labSNo IsNot Nothing Then labSNo.Text = TIMS.Get_DGSeqNo(sender, e)
    'End Sub

    Private Sub Datagrid12_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles Datagrid12.ItemDataBound
        If e Is Nothing Then Return
        Dim labSNo As Label = e.Item.FindControl("lab_SNo12")
        If labSNo IsNot Nothing Then labSNo.Text = TIMS.Get_DGSeqNo(sender, e)
    End Sub

    Private Sub Datagrid12b_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles Datagrid12b.ItemDataBound
        If e Is Nothing Then Return
        Dim labSNo As Label = e.Item.FindControl("lab_SNo12b")
        If labSNo IsNot Nothing Then labSNo.Text = TIMS.Get_DGSeqNo(sender, e)
    End Sub

    Private Sub Datagrid13_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles Datagrid13.ItemDataBound
        If e Is Nothing Then Return
        Dim labSNo As Label = e.Item.FindControl("lab_SNo13")
        If labSNo IsNot Nothing Then labSNo.Text = TIMS.Get_DGSeqNo(sender, e)
    End Sub

    Private Sub Datagrid14_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles Datagrid14.ItemDataBound
        If e Is Nothing Then Return
        Dim labSNo As Label = e.Item.FindControl("lab_SNo14")
        If labSNo IsNot Nothing Then labSNo.Text = TIMS.Get_DGSeqNo(sender, e)

        Const Cst_AppliedStatusM As Integer = 10
        Const Cst_AppliedStatus As Integer = 11

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                '申請補助金額 預算別 審核狀態 撥款狀態 訓練狀態 
                '審核狀態
                Dim StatusMMsg1 As String = ""
                Select Case Convert.ToString(drv("AppliedStatusM"))
                    Case cst_AppliedStatusM_NOINFO
                        StatusMMsg1 = cst_StatusMsg1_NOINFO '"審核通過" '"申請成功"
                    Case "Y"
                        StatusMMsg1 = cst_StatusMsg1_Y '"審核通過" '"申請成功"
                    Case "N"
                        StatusMMsg1 = cst_StatusMsg1_N '"審核不通過" '"申請失敗"
                    Case "R"
                        StatusMMsg1 = cst_StatusMsg1_R '"退件修正"
                    Case Else
                        'e.SumOfMoney
                        If Convert.ToString(drv("SumOfMoney")) <> "" AndAlso Val(drv("SumOfMoney")) > 0 Then
                            StatusMMsg1 = cst_StatusMsg1_S '"審核中" '"未審核"
                        End If
                        'If Convert.ToString(drv("SumOfMoney")) <> "" Then StatusMMsg1 = cst_StatusMsg1_S '"審核中" '"未審核"
                End Select
                e.Item.Cells(Cst_AppliedStatusM).Text = StatusMMsg1

                '撥款狀態
                If Convert.ToString(drv("AppliedStatus")) = "1" Then
                    e.Item.Cells(Cst_AppliedStatus).Text = cst_StatusMsg2_1 '"已撥款" '"申請成功"
                Else
                    Dim StatusMsg2 As String = ""
                    Select Case Convert.ToString(drv("AppliedStatusM"))
                        Case cst_AppliedStatusM_NOINFO
                            StatusMsg2 = cst_StatusMsg2_NOINFO '"審核通過" '"申請成功"
                        Case "Y" '審核通過
                            StatusMsg2 = cst_StatusMsg2_Y '"撥款中" '"申請中"
                        Case "N" '審核不通過
                            StatusMsg2 = cst_StatusMsg2_N '"不撥款" '"申請中"
                        Case "R" '退件修正
                            StatusMsg2 = cst_StatusMsg2_R '"未撥款" '"申請失敗"
                        Case Else '審核中
                            'StatusMsg2 = cst_StatusMsg2_X 'cst_StatusMsg2_R 'e.Item.Cells(Cst_AppliedStatus).Text = cst_StatusMsg2_R '"未撥款" '"申請失敗"
                            Select Case Convert.ToString(drv("SdStatus"))
                                Case "2", "3"
                                    StatusMsg2 = cst_StatusMsg2_X
                                    'e.Item.Cells(Cst_AppliedStatus).Text = cst_StatusMsg2_X '"不予補助"
                            End Select
                            'Case cst_AppliedStatusM_Y2 '審核通過 StatusMsg2 = cst_StatusMsg2_Y2 '"待撥款" "撥款中" '"申請中"
                    End Select
                    e.Item.Cells(Cst_AppliedStatus).Text = StatusMsg2
                End If

        End Select

    End Sub

End Class
