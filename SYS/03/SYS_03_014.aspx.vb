Public Class SYS_03_014
    Inherits AuthBasePage

    '異動Table : Auth_REndClass2
    'MakeListItem: 從Sql查詢結果集合的第1欄位(Value)、第2欄位(Text)

#Region "Function"

    '年度
    Function GetTPlanIDYears(ByVal PlanID As String, ByRef TPlanID As String, ByRef Years As String) As DataRow
        Dim sql As String = "SELECT * FROM id_Plan Where PlanID=@PlanID "
        Dim parms As Hashtable = New Hashtable()
        parms.Clear()
        parms.Add("PlanID", PlanID)
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)
        If Not dr Is Nothing Then
            TPlanID = dr("TPlanID")
            Years = dr("Years")
        End If
        Return dr
    End Function

    '計畫
    Sub Makeplanlist(ByRef ddlobj As DropDownList, ByVal Years As String, ByVal DistID As String)
        Dim sql As String = ""
        sql = "" & vbCrLf
        '順序有差1.ID 2.NAME
        sql += " select distinct a.PlanID, a.Years+b.Name+c.PlanName+a.seq PlanName " & vbCrLf
        sql += " ,a.DistID,a.TPlanID"
        sql += " from ID_Plan a " & vbCrLf
        sql += " JOIN ID_District b on a.DistID=b.DistID" & vbCrLf
        sql += " JOIN Key_Plan c on a.TPlanID=c.TPlanID" & vbCrLf
        sql += " JOIN Auth_AccRWPlan d on a.PlanID=d.PlanID" & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " and a.years = '" & Years & "' " & vbCrLf
        sql += " and a.DistID = '" & DistID & "'" & vbCrLf
        'TIMS.Cst_TPlanID28_2 
        sql += " and a.TPlanID in (" & TIMS.Cst_TPlanID28_2 & ") " & vbCrLf
        sql += " order by 2 " & vbCrLf

        ddlobj.Items.Clear()
        DbAccess.MakeListItem(ddlobj, sql, objconn)
        ddlobj.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    End Sub

    '機構
    Sub MakeddlOrgName(ByRef ddlobj As DropDownList)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT distinct oo.OrgID, oo.OrgName, b.OrgLevel, b.RID,b.DistID " & vbCrLf
        sql += "  FROM Auth_Relship b " & vbCrLf
        sql += "  JOIN Org_Orginfo oo on oo.OrgID =b.OrgID" & vbCrLf
        sql += "  WHERE 1=1" & vbCrLf
        sql += "  and b.PlanID = '0' " & vbCrLf
        sql += "  order by b.DistID,oo.OrgName,b.OrgLevel, b.RID " & vbCrLf

        ddlobj.Items.Clear()
        DbAccess.MakeListItem(ddlobj, sql, objconn)
        'ddlobj.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        ddlobj.Items.Insert(0, New ListItem("全部", ""))
    End Sub

    '帳號
    Sub MakeAccount(ByRef ddlobj As DropDownList,
                    ByVal PlanID As String, ByVal LID As String, ByVal RoleID As String, ByVal DistID As String, ByVal OrgID As String)
        'Dim TPlanID As String = ""
        'Dim Years As String = ""
        'Dim dr As DataRow = GetTPlanIDYears(PlanID, TPlanID, Years)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT distinct a.Account ,a.Name+'('+d.Name+')' sName" & vbCrLf
        sql += " ,a.RoleID,a.LID" & vbCrLf
        sql += " From Auth_Account a" & vbCrLf
        sql += " JOIN Auth_AccRWPlan b ON a.Account=b.Account" & vbCrLf
        sql += " JOIN Auth_Relship c ON b.RID=c.RID" & vbCrLf
        sql += " LEFT JOIN id_plan ip ON ip.PlanID =c.PlanID" & vbCrLf
        sql += " LEFT JOIN ID_Role d ON a.RoleID=d.RoleID" & vbCrLf
        sql += " Where 1=1" & vbCrLf
        sql += " and a.IsUsed='Y' and a.LID!='2'" & vbCrLf '有效帳號且排除委訓單位
        sql += " and a.LID>='" & LID & "' " & vbCrLf
        sql += " and a.RoleID>='" & RoleID & "' " & vbCrLf
        sql += " and b.PlanID = '" & PlanID & "' " & vbCrLf
        sql += " and c.DistID = '" & DistID & "' " & vbCrLf
        If OrgID <> "" Then
            sql += " and c.OrgID = '" & OrgID & "' " & vbCrLf
        End If
        sql += " order by a.RoleID, sName" & vbCrLf

        ddlobj.Items.Clear()
        DbAccess.MakeListItem(ddlobj, sql, objconn)
        ddlobj.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    End Sub

    '取得目前 已結訓班級使用授權檔 流水號 最大值
    Function Auth_REndClass2_MaxNo() As Integer
        Dim MaxNo As Integer = 1
        Dim sql As String = ""
        Dim dr As DataRow
        sql = "select max(RightID) max from Auth_REndClass2 "
        dr = DbAccess.GetOneRow(sql, objconn)
        If Not dr Is Nothing Then
            If Not IsDBNull(dr("max")) Then
                MaxNo = CInt(dr("max")) + 1
            End If
        End If
        Return MaxNo
    End Function

    '所選擇的開放功能值
    Function chkSelFunID() As String
        Dim selstr As String = ""
        selstr = ""
        For i As Int16 = 0 To cb_SelFunID.Items.Count - 1
            If cb_SelFunID.Items(i).Selected Then
                If selstr <> "" Then selstr += ","
                selstr += cb_SelFunID.Items(i).Value.ToString()
            End If
        Next
        Return selstr
    End Function

    '檢查帳號是否已賦于權限 (每個班級只限授予一個帳號 使用中的)
    Private Function chkNoRecordInAuth_Rend(ByVal Account As String, ByVal OCID As String) As Boolean
        Dim NoRecord As Boolean = True

        For i As Integer = 0 To cb_SelFunID.Items.Count - 1
            If cb_SelFunID.Items(i).Selected Then
                Dim dt As DataTable
                Dim sql As String = ""
                Dim selstr As String = ""
                selstr = cb_SelFunID.Items(i).Value.ToString()

                sql = ""
                sql += " select RightID,Years,OCID,account,UseAble,EndDate,FunID from Auth_RendClass2 "
                sql += " where UseAble='Y' "
                sql += " and OCID=@OCID " '20090302  改每個班級只限授予一個帳號

                Dim parms As Hashtable = New Hashtable()
                parms.Clear()
                parms.Add("OCID", OCID)
                dt = DbAccess.GetDataTable(sql, objconn, parms)
                If dt.Rows.Count > 0 Then
                    NoRecord = False
                    Exit For
                End If
            End If
        Next
        Return NoRecord
    End Function

    Function Chkdata(ByRef MsgStr As String) As Boolean
        Dim rst As Boolean = True
        Dim ErrCount As Int16 = 0

        MsgStr = ""
        If Me.Account.SelectedValue = "" Then
            MsgStr += "請選擇【帳號】!" & vbCrLf
            ErrCount = ErrCount + 1
        End If
        If Me.ReasonID.SelectedValue = "" Then
            MsgStr += "請選擇【補登資料原因】!" & vbCrLf
            ErrCount = ErrCount + 1
        End If
        If Me.Reason.Text = "" Then
            MsgStr += "請填寫【補登資料原因簡述】!" & vbCrLf
            ErrCount = ErrCount + 1
        End If
        If chkSelFunID() = "" Then
            MsgStr += "請選擇【開放功能】！" & vbCrLf
            ErrCount = ErrCount + 1
        End If
        If EndDate.Text = "" Then
            MsgStr += "請選擇【結束日期】！" & vbCrLf
            ErrCount = ErrCount + 1
        End If

        If ErrCount > 0 Then
            rst = False
        End If

        Return rst
    End Function

    Function chk_search(ByRef MsgStr As String) As Boolean
        Dim rst As Boolean = True
        Dim ErrCount As Integer = 0

        MsgStr = ""
        If Me.yearlist.SelectedValue = "" Then
            MsgStr += "請選擇【年度】!" & vbCrLf
            ErrCount = ErrCount + 1
        End If
        If Me.DistID.SelectedValue = "" Then
            MsgStr += "請選擇【轄區】!" & vbCrLf
            ErrCount = ErrCount + 1
        End If
        If Me.planlist.SelectedValue = "" Then
            MsgStr += "請選擇【訓練計畫】!" & vbCrLf
            ErrCount = ErrCount + 1
        End If

        If ErrCount > 0 Then
            rst = False
        End If

        Return rst
    End Function

    '取得開放授權使用功能
    Public Shared Function Get_FunIDReUse2(ByVal obj As ListControl, ByRef tConn As SqlConnection) As ListControl
        'Dim sql As String = ""
        ''限定：學員資料維護: 83
        'sql = ""
        'sql &= " SELECT a.Name +': '+a.FunID Name1,a.* FROM ID_Function a WHERE 1=1" & vbCrLf
        'sql &= " AND a.ReUse='Y'" & vbCrLf
        'sql &= " AND a.FunID IN (83)" & vbCrLf
        'sql &= " ORDER BY 1" & vbCrLf
        'Dim dt As New DataTable
        'Call TIMS.OpenDbConn(tConn)
        'Dim oCmd As New SqlCommand(sql, tConn)
        'With oCmd
        '    .Parameters.Clear()
        '    dt.Load(.ExecuteReader())
        'End With
        obj = TIMS.Get_FunIDReUse(obj, tConn, "83")
        Return obj
    End Function

#End Region

    '2011 功能按鈕權限控管參數 ---------------------Start
    'Dim blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt As Boolean '新增'修改'刪除'查詢'列印
    '2011 功能按鈕權限控管參數 ---------------------End

    '異動Table : Auth_REndClass2
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        'PageControler1 = Me.FindControl("PageControler1")
        PageControler1.PageDataGrid = DG_ClassInfo

        'Dim dr As DataRow
        'Dim sql As String
        If Not IsPostBack Then
            msg.Text = ""
            ReasonID = TIMS.Get_ReasonID(ReasonID, objconn)
            yearlist = TIMS.GetSyear(yearlist, 0, 0, False)
            DistID = TIMS.Get_DistID(DistID)
            '可使用的補登功能 存在 ID_Function (select * from ID_Function WHERE ReUse='Y')
            cb_SelFunID = Get_FunIDReUse2(cb_SelFunID, objconn)
            Common.SetListItem(yearlist, Now.Year)

            Reason_tr.Visible = False
            Account_tr.Visible = False
            Me.trOrgName.Visible = False
            '-----------------20090226 andy add
            Fun_tr.Visible = False
            '----------------
            PageControler1.Visible = False
        End If

        'rt_search.Attributes("onclick") = "javascript:return search()"

        '檢查帳號的功能權限-----------------------------------Start
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        'Dim MyValue As String = ""
        'check_Sech.Value = TIMS.GetMyValue(Session(TIMS.Cst_FunAuth), "Sech") '查詢
        'check_add.Value = TIMS.GetMyValue(Session(TIMS.Cst_FunAuth), "Adds") '新增資料權
        'check_mod.Value = TIMS.GetMyValue(Session(TIMS.Cst_FunAuth), "Mod") '修改資料權
        'check_del.Value = TIMS.GetMyValue(Session(TIMS.Cst_FunAuth), "Del") '刪除資料權

        'If check_add.Value = "1" OrElse check_mod.Value = "1" OrElse check_del.Value = "1" OrElse check_Sech.Value = "1" Then check_Sech.Value = "1"
        '順序有差別
        rt_search.Enabled = True '可查詢
        'If check_Sech.Value = "0" Then rt_search.Enabled = False '不可查詢
        '檢查帳號的功能權限-----------------------------------End
    End Sub

    Private Sub yearlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yearlist.SelectedIndexChanged
        '若沒有選擇轄區帶入使用者登入轄區
        If Me.DistID.SelectedValue = "" Then
            Common.SetListItem(Me.DistID, sm.UserInfo.DistID)
        End If
        '計畫
        Call Makeplanlist(planlist, Me.yearlist.SelectedValue, Me.DistID.SelectedValue)

        TPanel.Visible = False
        DG_ClassInfo.Visible = False 'TPanel@DG_ClassInfo
        PageControler1.Visible = False 'TPanel@PageControler1
        Reason_tr.Visible = False
        Account_tr.Visible = False
        Me.trOrgName.Visible = False
        Label2.Visible = False
        '------------------20090226 andy add
        Fun_tr.Visible = False
        '----------------
        'msg.Text = "查無資料!!"
    End Sub

    Private Sub DistID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DistID.SelectedIndexChanged
        '計畫
        Call Makeplanlist(planlist, Me.yearlist.SelectedValue, Me.DistID.SelectedValue)

        TPanel.Visible = False
        DG_ClassInfo.Visible = False 'TPanel@DG_ClassInfo
        PageControler1.Visible = False 'TPanel@PageControler1
        Reason_tr.Visible = False
        Account_tr.Visible = False
        Me.trOrgName.Visible = False
        Label2.Visible = False
        '------------------20090226 andy add
        Fun_tr.Visible = False
        '----------------
        'msg.Text = "查無資料!!"
    End Sub

    Function ShowDG_ClassInfo(ByVal dt As DataTable) As DataTable
        'Dim IsTPlan28 As Boolean = False
        'Dim vsTPlanID As String = ""
        'IsTPlan28 = False
        'vsTPlanID = TIMS.GetTPlanID(Me.planlist.SelectedValue)
        ''產業人才投資方案
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(vsTPlanID) > -1 Then
        '    IsTPlan28 = True
        'End If
        Dim parms As Hashtable = New Hashtable()
        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr += " Select a.Years" & vbCrLf
        sqlstr += " ,a.CyclType" & vbCrLf
        sqlstr += " ,a.ClassNum" & vbCrLf
        sqlstr += " ,b.ClassID" & vbCrLf
        sqlstr += " ,a.PlanID" & vbCrLf
        sqlstr += " ,a.OCID" & vbCrLf
        sqlstr += " ,e.OrgName" & vbCrLf
        sqlstr &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSCNAME" & vbCrLf
        sqlstr += " ,g.TrainName" & vbCrLf
        sqlstr += " ,a.STDate" & vbCrLf
        sqlstr += " ,a.FTDate" & vbCrLf
        sqlstr += " ,a.RID" & vbCrLf

#Region "oracle ver no use"
        'sqlstr += " ,dbo.NVL(CONVERT(varchar, h.RightID),'XX') RightID" & vbCrLf
        'sqlstr += " ,dbo.NVL(h.NAME,' ')  NAME" & vbCrLf
        'sqlstr += " ,dbo.NVL(h.ACCOUNT,' ')  ACCOUNT" & vbCrLf
        'sqlstr += " ,dbo.DECODE(dbo.NVL(h.ACCOUNT,'0'),'0','0','1') Acnt" & vbCrLf
#End Region
        sqlstr += " ,ISNULL(CONVERT(varchar, H.RIGHTID),'XX') RIGHTID " & vbCrLf
        sqlstr += " ,ISNULL(H.NAME,' ')  NAME " & vbCrLf
        sqlstr += " ,ISNULL(H.ACCOUNT,' ')  ACCOUNT " & vbCrLf
        sqlstr += " ,DBO.DECODE(ISNULL(H.ACCOUNT,'0'),'0','0','1') Acnt " & vbCrLf

        sqlstr += " ,h.EndDate" & vbCrLf
        sqlstr += " ,h.Temp1" & vbCrLf
        sqlstr += " ,a.Years + '0' + b.ClassID + a.CyclType  ClassID2" & vbCrLf
        sqlstr += " From Class_ClassInfo a" & vbCrLf
        sqlstr += " join id_plan ip on ip.PlanID =a.PlanID" & vbCrLf
        sqlstr += " join ID_Class b on a.CLSID = b.CLSID" & vbCrLf
        sqlstr += " JOIN ID_District c ON b.DistID = c.DistID" & vbCrLf
        sqlstr += " JOIN Auth_Relship d on a.RID  = d.RID" & vbCrLf
        sqlstr += " JOIN Org_OrgInfo e on d.OrgID = e.OrgID" & vbCrLf
        sqlstr += " LEFT JOIN Key_TrainType g   on a.TMID  = g.TMID" & vbCrLf
        sqlstr += " LEFT JOIN (" & vbCrLf
        sqlstr += "     SELECT h1.RightID,h1.OCID,h2.ACCOUNT,h2.NAME,h1.EndDate " & vbCrLf
        sqlstr += "     ,d1.Name+';開放FunID: '+h1.FunID  Temp1" & vbCrLf
        sqlstr += " 	FROM Auth_REndClass2 h1" & vbCrLf
        sqlstr += " 	join Auth_Account h2 ON h1.ACCOUNT = h2.ACCOUNT" & vbCrLf
        sqlstr += " 	left join ID_keyinReason d1 on h1.ReasonID=d1.ReasonID" & vbCrLf
        sqlstr += " 	where h1.UseAble = 'Y'" & vbCrLf
        sqlstr += " ) h ON a.OCID = h.OCID" & vbCrLf
        sqlstr += " Where 1=1 " & vbCrLf
        '20090617 andy  edit
        '--------------------
        sqlstr += "  and a.IsSuccess='Y'" & vbCrLf '是否轉入成功
        sqlstr += "  and a.NotOpen='N' " & vbCrLf  '不開班
        sqlstr += "  and ip.Years = @Years " & vbCrLf
        sqlstr += "  and ip.DistID = @DistID " & vbCrLf

        parms.Add("@Years", Me.yearlist.SelectedValue)
        parms.Add("@DistID", Me.DistID.SelectedValue)

        If Me.planlist.SelectedValue <> "" Then
            sqlstr += " and ip.PlanID = @PlanID " & vbCrLf
            parms.Add("PlanID", Me.planlist.SelectedValue)
        End If
        Select Case ClassRound.SelectedIndex
            Case 0 '已結訓
                'sqlstr += " and a.FTDate+1 <= getdate()" & vbCrLf
                sqlstr += " and dbo.TRUNC_DATETIME(a.FTDate) <= dbo.TRUNC_DATETIME(getdate())" & vbCrLf
            Case 1 '未結訓
                'sqlstr += " and a.FTDate+1 > getdate()" & vbCrLf
                sqlstr += " and dbo.TRUNC_DATETIME(a.FTDate) > dbo.TRUNC_DATETIME(getdate())" & vbCrLf
        End Select
        '-------------------

        If Me.ClassName.Text <> "" Then
            sqlstr += " and a.ClassCName like @ClassCName " & vbCrLf
            parms.Add("ClassCName", "%" & Me.ClassName.Text & "%")
        End If
        If Me.start_date.Text <> "" Then
            sqlstr += "and a.STDate >= @STDate1 " & vbCrLf
            parms.Add("STDate1", Me.start_date.Text)
        End If
        If Me.end_date.Text <> "" Then
            sqlstr += " and a.STDate <= @STDate2 " & vbCrLf
            parms.Add("STDate2", Me.end_date.Text)
        End If
        If Me.CyclType.Text <> "" Then
            If IsNumeric(Me.CyclType.Text) Then
                'If Int(Me.CyclType.Text) < 10 Then
                '    sqlstr += " and a.CyclType = '0" & Int(Me.CyclType.Text) & "'" & vbCrLf
                'Else
                '    sqlstr += " and a.CyclType = '" & Me.CyclType.Text & "'" & vbCrLf
                'End If
                sqlstr += " and a.CyclType = @CyclType " & vbCrLf
                parms.Add("CyclType", IIf(Int(Me.CyclType.Text) < 10, "0", "") & Int(Me.CyclType.Text))
            End If
        End If

        dt = DbAccess.GetDataTable(sqlstr, objconn, parms)
        Return dt
    End Function

    Sub dt_search()
        TPanel.Visible = False
        DG_ClassInfo.Visible = False 'TPanel@DG_ClassInfo
        PageControler1.Visible = False 'TPanel@PageControler1
        Reason_tr.Visible = False
        Account_tr.Visible = False
        Me.trOrgName.Visible = False
        Label2.Visible = False
        '------------------20090226 andy add
        Fun_tr.Visible = False
        '----------------
        msg.Text = "查無資料!!"

        Dim MsgStr As String = ""
        MsgStr = ""
        If Not chk_search(MsgStr) Then
            Common.MessageBox(Me, MsgStr)
            Exit Sub
        End If

        Dim vsOrgName As String = ""
        Dim vsAccount As String = ""
        vsOrgName = ""
        vsAccount = ""
        If Me.ddlOrgName.SelectedValue <> "" Then
            vsOrgName = Me.ddlOrgName.SelectedValue
        End If
        If Me.Account.SelectedValue <> "" Then
            vsAccount = Me.Account.SelectedValue
        End If
        '取得授權單位。
        'Call MakeddlOrgName(Me.ddlOrgName,  Me.planlist.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, Me.DistID.SelectedValue)
        Call MakeddlOrgName(Me.ddlOrgName)
        '帳號
        Call MakeAccount(Me.Account, Me.planlist.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, Me.DistID.SelectedValue, Me.ddlOrgName.SelectedValue)
        'Call MakeAccount(Me.Account, Me.planlist.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, Me.DistID.SelectedValue)
        If vsOrgName <> "" Then
            Common.SetListItem(ddlOrgName, vsOrgName)
        End If
        If vsAccount <> "" Then
            Common.SetListItem(Account, vsAccount)
        End If

        Dim dt As DataTable = Nothing
        dt = ShowDG_ClassInfo(dt)

        TPanel.Visible = False
        DG_ClassInfo.Visible = False 'TPanel@DG_ClassInfo
        PageControler1.Visible = False 'TPanel@PageControler1
        Reason_tr.Visible = False
        Account_tr.Visible = False
        Me.trOrgName.Visible = False
        Label2.Visible = False
        '------------------20090226 andy add
        Fun_tr.Visible = False
        '----------------
        msg.Text = "查無資料!!"

        If dt.Rows.Count > 0 Then
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()

            TPanel.Visible = True
            DG_ClassInfo.Visible = True 'TPanel@DG_ClassInfo
            PageControler1.Visible = True 'TPanel@PageControler1
            Reason_tr.Visible = True
            Account_tr.Visible = True

            Me.trOrgName.Visible = True
            Label2.Visible = True
            '------------------20090226 andy add
            Fun_tr.Visible = True
            '----------------
            msg.Text = ""
        End If
    End Sub

    Private Sub DG_ClassInfo_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_ClassInfo.ItemCommand

        'Const Cst_選擇 As Integer = 0
        'Const Cst_OrgName As Integer = 1
        'Const Cst_ClassID2 As Integer = 2
        'Const Cst_STDate As Integer = 3 '開訓日期
        Const Cst_FTDate As Integer = 4 '結訓日期
        'Const Cst_ClassCName As Integer = 5
        'Const Cst_TrainName As Integer = 6
        Const Cst_OCID As Integer = 7 '班級代碼
        Const Cst_RightID As Integer = 8 '已結訓班級使用授權檔
        'Const Cst_Name As Integer = 9 '已授權給(授權帳號:姓名)

        Dim sql As String = ""
        Dim MsgStr As String = ""
        Dim strOCID As String = e.Item.Cells(Cst_OCID).Text
        Dim strRightID As String = e.Item.Cells(Cst_RightID).Text
        Dim strFDate As String = e.Item.Cells(Cst_FTDate).Text

        Select Case e.CommandName
            Case "Add"
                '20090617 andy  edit  只有「e網報名審核-匯入」功能不管結訓日期，開放其它功能要判斷該班結訓日期
                '---------------------  
                'Dim i As Integer = 0

                For i As Integer = 0 To cb_SelFunID.Items.Count - 1
                    If cb_SelFunID.Items(i).Selected = True Then
                        Dim flagCheckGo As Boolean = True  '需要檢核 判斷該班結訓日期
                        Select Case cb_SelFunID.Items(i).Value
                            Case "262", "83"
                                'e網報名審核: 262
                                '學員資料維護: 83
                                flagCheckGo = False '排除檢核
                        End Select

                        If flagCheckGo Then
                            If strFDate <> "" Then
                                If CDate(CDate(strFDate).ToString("yyyy-MM-dd")) > CDate(Date.Now.ToString("yyyy-MM-dd")) Then
                                    Common.MessageBox(Me, "此班尚未結訓！")
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                Next
                '---------------------
                '20090226 andy  edit
                '-------------------
                MsgStr = ""
                If Chkdata(MsgStr) = False Then
                    Common.MessageBox(Me, MsgStr)
                    Exit Sub
                End If
                If chkNoRecordInAuth_Rend(Account.SelectedValue.ToString(), strOCID) = False Then
                    Common.MessageBox(Me, "同一班級只限授權一個帳號！")
                    Exit Sub
                End If
                '-------------------
                sql = ""
                sql &= " INSERT INTO Auth_REndClass2 "
                sql += " (RightID,Years,PlanID,DistID,OCID,Account,CreateDate,UseAble,ModifyAcct,ModifyDate,Reason,ReasonID,FunID,EndDate) "
                sql &= " values"
                sql += " (@RightID,@Years,@PlanID,@DistID,@OCID,@Account,getdate(),@UseAble,@ModifyAcct,getdate(),@Reason,@ReasonID,@FunID,@EndDate) "
                Call TIMS.OpenDbConn(objconn)
                Dim oCmd As New SqlCommand(sql, objconn)
                With oCmd
                    .Parameters.Clear()
                    .Parameters.Add("RightID", SqlDbType.Int).Value = Auth_REndClass2_MaxNo()
                    .Parameters.Add("Years", SqlDbType.VarChar).Value = Me.yearlist.SelectedValue
                    .Parameters.Add("PlanID", SqlDbType.VarChar).Value = Me.planlist.SelectedValue
                    .Parameters.Add("DistID", SqlDbType.VarChar).Value = Me.DistID.SelectedValue
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = strOCID
                    .Parameters.Add("Account", SqlDbType.VarChar).Value = Me.Account.SelectedValue
                    '.Parameters.Add("CreateDate", SqlDbType.VarChar).Value = getdate()
                    .Parameters.Add("UseAble", SqlDbType.VarChar).Value = "Y" 'UseAble
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    '.Parameters.Add("ModifyDate", SqlDbType.VarChar).Value = getdate()
                    .Parameters.Add("Reason", SqlDbType.VarChar).Value = Me.Reason.Text
                    .Parameters.Add("ReasonID", SqlDbType.VarChar).Value = Me.ReasonID.SelectedValue
                    .Parameters.Add("FunID", SqlDbType.VarChar).Value = chkSelFunID()
                    .Parameters.Add("EndDate", SqlDbType.DateTime).Value = TIMS.Cdate2(Me.EndDate.Text)
                    '.ExecuteNonQuery()
                    ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                    DbAccess.ExecuteNonQuery(oCmd.CommandText, objconn, oCmd.Parameters)
                End With
                Common.MessageBox(Me, "新增成功")

                Call dt_search()
            Case "Upd"
                '20090617 andy  edit  只有「e網報名審核-匯入」功能不管結訓日期，開放其它功能要判斷該班結訓日期
                '---------------------  
                'Dim i As Integer = 0
                For i As Integer = 0 To cb_SelFunID.Items.Count - 1
                    If cb_SelFunID.Items(i).Selected = True Then
                        Dim flagCheckGo As Boolean = True  '需要檢核 判斷該班結訓日期
                        Select Case cb_SelFunID.Items(i).Value
                            Case "262", "83"
                                'e網報名審核: 262
                                '學員資料維護: 83
                                flagCheckGo = False '排除檢核
                        End Select
                        If flagCheckGo Then
                            If strFDate <> "" Then
                                If CDate(CDate(strFDate).ToString("yyyy-MM-dd")) > CDate(Date.Now.ToString("yyyy-MM-dd")) Then
                                    Common.MessageBox(Me, "此班尚未結訓！")
                                    Exit Sub
                                End If
                            End If
                        End If

                    End If
                Next
                '---------------------
                '20090226 andy  edit
                '-------------------
                If strRightID = "XX" Then '已結訓班級使用授權檔
                    Common.MessageBox(Me, "目前尚未針對本班級授權,無法進行修改動作！")
                    Exit Sub
                End If
                MsgStr = ""
                If Chkdata(MsgStr) = False Then
                    Common.MessageBox(Me, MsgStr)
                    Exit Sub
                End If
                '---------------
                sql = "" & vbCrLf
                sql &= " UPDATE Auth_REndClass2 " & vbCrLf
                sql += " SET Account = @Account" & vbCrLf
                sql += " ,Reason = @Reason" & vbCrLf
                sql += " ,ReasonID = @ReasonID" & vbCrLf
                sql += " ,ModifyAcct = @ModifyAcct" & vbCrLf
                sql += " ,ModifyDate = getdate()" & vbCrLf
                sql += " ,FunID = @FunID" & vbCrLf
                sql += " ,EndDate = @EndDate" & vbCrLf
                sql += " WHERE RightID = @RightID " & vbCrLf
                sql += " And OCID = @OCID" & vbCrLf

                Call TIMS.OpenDbConn(objconn)
                Dim oCmd As New SqlCommand(sql, objconn)
                With oCmd
                    .Parameters.Clear()
                    '.Parameters.Add("Years", SqlDbType.VarChar).Value = Me.yearlist.SelectedValue
                    '.Parameters.Add("PlanID", SqlDbType.VarChar).Value = Me.planlist.SelectedValue
                    '.Parameters.Add("DistID", SqlDbType.VarChar).Value = Me.DistID.SelectedValue
                    ''.Parameters.Add("CreateDate", SqlDbType.VarChar).Value = getdate()
                    '.Parameters.Add("UseAble", SqlDbType.VarChar).Value = "Y" 'UseAble

                    .Parameters.Add("Account", SqlDbType.VarChar).Value = Me.Account.SelectedValue
                    .Parameters.Add("Reason", SqlDbType.VarChar).Value = Me.Reason.Text
                    .Parameters.Add("ReasonID", SqlDbType.VarChar).Value = Me.ReasonID.SelectedValue
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    '.Parameters.Add("ModifyDate", SqlDbType.VarChar).Value = getdate()
                    .Parameters.Add("FunID", SqlDbType.VarChar).Value = chkSelFunID()
                    .Parameters.Add("EndDate", SqlDbType.DateTime).Value = TIMS.Cdate2(Me.EndDate.Text)
                    .Parameters.Add("RightID", SqlDbType.Int).Value = Val(strRightID)
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = strOCID
                    '.ExecuteNonQuery()

                    ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                    DbAccess.ExecuteNonQuery(oCmd.CommandText, objconn, oCmd.Parameters)
                End With
                Common.MessageBox(Me, "修改成功")

                Call dt_search()
            Case "Del"

                If strRightID = "XX" Then
                    Common.MessageBox(Me, "目前尚未針對本班級授權,無法進行刪除動作！")
                    Exit Sub
                End If
                sql = ""
                sql &= " UPDATE Auth_REndClass2 " & vbCrLf
                sql += " SET UseAble = 'N'" & vbCrLf
                sql += " ,ModifyAcct =@ModifyAcct" & vbCrLf
                sql += " ,ModifyDate = getdate() " & vbCrLf
                sql += " WHERE RightID = @RightID" & vbCrLf
                sql += " And OCID = @OCID" & vbCrLf
                'DbAccess.ExecuteNonQuery(sql, objconn)
                Call TIMS.OpenDbConn(objconn)
                Dim oCmd As New SqlCommand(sql, objconn)
                With oCmd
                    .Parameters.Clear()
                    '.Parameters.Add("Years", SqlDbType.VarChar).Value = Me.yearlist.SelectedValue
                    '.Parameters.Add("PlanID", SqlDbType.VarChar).Value = Me.planlist.SelectedValue
                    '.Parameters.Add("DistID", SqlDbType.VarChar).Value = Me.DistID.SelectedValue
                    ''.Parameters.Add("CreateDate", SqlDbType.VarChar).Value = getdate()
                    '.Parameters.Add("UseAble", SqlDbType.VarChar).Value = "Y" 'UseAble

                    '.Parameters.Add("Account", SqlDbType.VarChar).Value = Me.Account.SelectedValue
                    '.Parameters.Add("Reason", SqlDbType.VarChar).Value = Me.Reason.Text
                    '.Parameters.Add("ReasonID", SqlDbType.VarChar).Value = Me.ReasonID.SelectedValue
                    ''.Parameters.Add("ModifyDate", SqlDbType.VarChar).Value = getdate()
                    '.Parameters.Add("FunID", SqlDbType.VarChar).Value = chkSelFunID()
                    '.Parameters.Add("EndDate", SqlDbType.DateTime).Value = TIMS.cdate2(Me.EndDate.Text)

                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .Parameters.Add("RightID", SqlDbType.Int).Value = Val(strRightID)
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = strOCID
                    '.ExecuteNonQuery()
                    ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                    DbAccess.ExecuteNonQuery(oCmd.CommandText, objconn, oCmd.Parameters)
                End With
                Common.MessageBox(Me, "刪除成功")

                Call dt_search()
        End Select

    End Sub

    Private Sub DG_ClassInfo_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_ClassInfo.ItemDataBound
        Dim drv As DataRowView = e.Item.DataItem

        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim but1 As LinkButton = e.Item.FindControl("but1") '新增
                Dim but2 As LinkButton = e.Item.FindControl("but2") '修改
                Dim but3 As LinkButton = e.Item.FindControl("but3") '刪除

                'Dim but4 As Button = e.Item.FindControl("but4") '查看
                'but4.CommandArgument = Convert.ToString(drv("RightID"))
                If Convert.ToString(drv("Temp1")) <> "" Then
                    TIMS.Tooltip(e.Item, Convert.ToString(drv("Temp1")))
                    TIMS.Tooltip(but1, Convert.ToString(drv("Temp1")))
                    TIMS.Tooltip(but2, Convert.ToString(drv("Temp1")))
                    TIMS.Tooltip(but3, Convert.ToString(drv("Temp1")))
                End If

                but3.Attributes("onclick") = "return confirm('確定要刪除此授權?');"

        End Select
    End Sub

    Private Sub rt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rt_search.Click
        dt_search()
    End Sub

    Private Sub planlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles planlist.SelectedIndexChanged
        dt_search()
    End Sub

    Private Sub ddlOrgName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddlOrgName.SelectedIndexChanged
        '帳號
        Call MakeAccount(Me.Account, Me.planlist.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, Me.DistID.SelectedValue, Me.ddlOrgName.SelectedValue)
        'Call MakeAccount(Me.Account, Me.planlist.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, Me.DistID.SelectedValue)
    End Sub

End Class