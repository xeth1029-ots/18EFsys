Partial Class SYS_06_009
    Inherits AuthBasePage

    '異動Table : Auth_REndClass
    'MakeListItem: 從Sql查詢結果集合的第1欄位(Value)、第2欄位(Text)

#Region "Function"

    '年度
    Function GetTPlanIDYears(ByVal PlanID As String, ByRef TPlanID As String, ByRef Years As String) As DataRow
        Dim sql As String = "SELECT * FROM id_Plan Where PlanID ='" & PlanID & "'"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        If Not dr Is Nothing Then
            TPlanID = dr("TPlanID")
            Years = dr("Years")
        End If
        Return dr
    End Function

    '計畫顯示。
    Sub Makeplanlist(ByRef ddlobj As DropDownList, ByVal Years As String, ByVal DistID As String)
        Dim Sql As String = ""
        Sql = "" & vbCrLf
        Sql += " select distinct a.PlanID, a.Years+b.Name+c.PlanName+a.seq PlanName, a.DistID " & vbCrLf
        Sql += " from ID_Plan a " & vbCrLf
        Sql += " JOIN ID_District b on a.DistID=b.DistID" & vbCrLf
        Sql += " JOIN Key_Plan c on a.TPlanID=c.TPlanID" & vbCrLf
        Sql += " JOIN Auth_AccRWPlan d on a.PlanID=d.PlanID" & vbCrLf
        Sql += " where 1=1" & vbCrLf
        Sql += " and a.Years = '" & Years & "' " & vbCrLf
        Sql += " and a.DistID = '" & DistID & "'" & vbCrLf
        Sql += " order by 2 " & vbCrLf

        ddlobj.Items.Clear()
        DbAccess.MakeListItem(ddlobj, Sql, objconn)
        ddlobj.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    End Sub

    '機構
    Sub MakeddlOrgName(ByRef ddlobj As DropDownList, ByVal PlanID As String, ByVal LID As String, ByVal RoleID As String, ByVal DistID As String)
        'Dim TPlanID As String = ""
        'Dim Years As String = ""
        'Dim dr As DataRow = GetTPlanIDYears(PlanID, TPlanID, Years)

        Dim Sql As String = ""
        Sql = "" & vbCrLf
        Sql += " SELECT distinct oo.OrgID, oo.OrgName, c.OrgLevel, c.RID " & vbCrLf
        Sql += " From Auth_Account a " & vbCrLf
        Sql += " JOIN Auth_AccRWPlan b ON a.Account=b.Account" & vbCrLf
        Sql += " JOIN Auth_Relship c ON b.RID=c.RID" & vbCrLf
        Sql += " JOIN Org_Orginfo oo on oo.OrgID =c.OrgID" & vbCrLf
        Sql += " LEFT JOIN id_plan ip ON ip.PlanID =c.PlanID " & vbCrLf
        Sql += " WHERE 1=1" & vbCrLf
        Sql += " and a.IsUsed='Y' " & vbCrLf
        Sql += " and a.LID>='" & LID & "' " & vbCrLf
        Sql += " and a.RoleID>='" & RoleID & "' " & vbCrLf
        Sql += " and b.PlanID = '" & PlanID & "' " & vbCrLf
        Sql += " and c.DistID = '" & DistID & "' " & vbCrLf
        'Sql += " and ip.TPlanID = '" & TPlanID & "' " & vbCrLf
        'Sql += " and ip.Years = '" & Years & "' " & vbCrLf
        Sql += " order by oo.OrgName,c.OrgLevel, c.RID " & vbCrLf

        ddlobj.Items.Clear()
        DbAccess.MakeListItem(ddlobj, Sql, objconn)
        'ddlobj.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        ddlobj.Items.Insert(0, New ListItem("全部", ""))
    End Sub

    '帳號
    Sub MakeAccount(ByRef ddlobj As DropDownList, ByVal PlanID As String, ByVal LID As String, ByVal RoleID As String, ByVal DistID As String, ByVal OrgID As String)
        'Dim TPlanID As String = ""
        'Dim Years As String = ""
        'Dim dr As DataRow = GetTPlanIDYears(PlanID, TPlanID, Years)

        Dim Sql As String = ""
        Sql = "" & vbCrLf
        Sql += " SELECT distinct a.Account ,a.Name+'('+d.Name+')' sName" & vbCrLf
        Sql += " ,a.RoleID,a.LID" & vbCrLf
        Sql += " From Auth_Account a" & vbCrLf
        Sql += " JOIN Auth_AccRWPlan b ON a.Account=b.Account" & vbCrLf
        Sql += " JOIN Auth_Relship c ON b.RID=c.RID" & vbCrLf
        Sql += " LEFT JOIN id_plan ip ON ip.PlanID =c.PlanID" & vbCrLf
        Sql += " LEFT JOIN ID_Role d ON a.RoleID=d.RoleID" & vbCrLf
        Sql += " Where 1=1" & vbCrLf
        Sql += " and a.IsUsed='Y' " & vbCrLf
        Sql += " and a.LID>='" & LID & "' " & vbCrLf
        Sql += " and a.RoleID>='" & RoleID & "' " & vbCrLf
        Sql += " and b.PlanID = '" & PlanID & "' " & vbCrLf
        Sql += " and c.DistID = '" & DistID & "' " & vbCrLf
        'Sql += " and ip.TPlanID = '" & TPlanID & "' " & vbCrLf
        'Sql += " and ip.Years = '" & Years & "' " & vbCrLf
        If OrgID <> "" Then
            Sql += " and c.OrgID = '" & OrgID & "' " & vbCrLf
        End If
        Sql += " order by a.RoleID, sName" & vbCrLf

        ddlobj.Items.Clear()
        DbAccess.MakeListItem(ddlobj, Sql, objconn)
        ddlobj.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    End Sub

    '取得目前 已結訓班級使用授權檔 流水號 最大值
    Function Auth_REndClass_MaxNo() As Integer
        Dim MaxNo As Integer = 1
        Dim sql As String = "select max(RightID) max from Auth_REndClass "
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
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

    '檢查帳號是否已賦于權限
    Private Function chkNoRecordInAuth_Rend(ByVal Account As String, ByVal OCID As String) As Boolean
        Dim NoRecord As Boolean = True

        For i As Integer = 0 To cb_SelFunID.Items.Count - 1
            If cb_SelFunID.Items(i).Selected Then
                Dim dt As DataTable
                Dim sql As String = ""
                Dim selstr As String = ""
                selstr = cb_SelFunID.Items(i).Value.ToString()

                sql = ""
                sql += " select RightID,Years,OCID,account,UseAble,EndDate,FunID from Auth_RendClass "
                sql += " where 1=1"
                sql += " and UseAble='Y' "
                sql += " and OCID='" & OCID & "'" '20090302  改每個班級只限一授于一個帳號
                dt = DbAccess.GetDataTable(sql, objconn)
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
        'If Me.ReasonID.SelectedValue = "" Then
        '    MsgStr += "請選擇【補登資料原因】!" & vbCrLf
        '    ErrCount = ErrCount + 1
        'End If
        'If Me.Reason.Text = "" Then
        '    MsgStr += "請填寫【補登資料原因簡述】!" & vbCrLf
        '    ErrCount = ErrCount + 1
        'End If
        'If chkSelFunID() = "" Then
        '    MsgStr += "請選擇【開放功能】！" & vbCrLf
        '    ErrCount = ErrCount + 1
        'End If
        'If EndDate.Text = "" Then
        '    MsgStr += "請選擇【結束日期】！" & vbCrLf
        '    ErrCount = ErrCount + 1
        'End If

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

#End Region

    '異動Table : Auth_REndClass
    'Dim FunDr As DataRow
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        If sm.UserInfo.FunDt Is Nothing Then
            Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        End If

        If Not IsPostBack Then
            msg.Text = ""

            yearlist = TIMS.GetSyear(yearlist, 0, 0, False)
            DistID = TIMS.Get_DistID(DistID)
            '可使用的補登功能 存在 ID_Function (select * from ID_Function WHERE ReUse='Y')
            cb_SelFunID = TIMS.Get_FunIDReUse(cb_SelFunID, objconn, "")
            Common.SetListItem(yearlist, Now.Year)

            Account_tr.Visible = False
            Me.trOrgName.Visible = False
            '-----------------20090226 andy add
            Fun_tr.Visible = False
            '----------------
            PageControler1.Visible = False
        End If

        'rt_search.Attributes("onclick") = "javascript:return search()"

        '檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.RoleID <> 0 Then
        '    If sm.UserInfo.FunDt Is Nothing Then
        '        Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '        Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        '    Else
        '        Dim FunDt As DataTable = sm.UserInfo.FunDt
        '        Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '        If FunDrArray.Length = 0 Then
        '            Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '        Else
        '            FunDr = FunDrArray(0)

        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End
    End Sub

    Private Sub yearlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yearlist.SelectedIndexChanged
        '若沒有選擇轄區帶入使用者登入轄區
        If Me.DistID.SelectedValue = "" Then
            Common.SetListItem(Me.DistID, sm.UserInfo.DistID)
        End If
        '計畫
        Call Makeplanlist(planlist, Me.yearlist.SelectedValue, Me.DistID.SelectedValue)

        'TPanel.Visible = False
        'DG_ClassInfo.Visible = False 'TPanel@DG_ClassInfo
        'PageControler1.Visible = False 'TPanel@PageControler1
        'Reason_tr.Visible = False
        'Account_tr.Visible = False
        'Me.trOrgName.Visible = False
        'Label2.Visible = False
        '------------------20090226 andy add
        Fun_tr.Visible = False
        '----------------
        'msg.Text = "查無資料!!"
    End Sub

    Private Sub DistID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DistID.SelectedIndexChanged
        '計畫
        Call Makeplanlist(planlist, Me.yearlist.SelectedValue, Me.DistID.SelectedValue)
        'MakeddlOrgName(Me.ddlOrgName, Me.planlist.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, Me.DistID.SelectedValue)
        'MakeAccount(Me.Account, Me.planlist.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, Me.DistID.SelectedValue, Me.ddlOrgName.SelectedValue)
        'TPanel.Visible = False
        'DG_ClassInfo.Visible = False 'TPanel@DG_ClassInfo
        'PageControler1.Visible = False 'TPanel@PageControler1
        'Reason_tr.Visible = False
        'Account_tr.Visible = False
        'Me.trOrgName.Visible = False
        'Label2.Visible = False
        '------------------20090226 andy add
        Fun_tr.Visible = False
        '----------------
        'msg.Text = "查無資料!!"
    End Sub

    'SQL
    Function ShowDG_ClassInfo(ByVal dt As DataTable) As DataTable
        'Dim IsTPlan28 As Boolean = False
        'Dim vsTPlanID As String = ""
        'IsTPlan28 = False
        'vsTPlanID = TIMS.GetTPlanID(Me.planlist.SelectedValue)
        ''產業人才投資方案
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(vsTPlanID) > -1 Then
        '    IsTPlan28 = True
        'End If
        Dim sSql As String = ""
        sSql &= " Select cc.Years" & vbCrLf
        sSql &= " ,cc.CyclType" & vbCrLf
        sSql &= " ,cc.ClassID" & vbCrLf
        sSql &= " ,cc.PlanID,cc.OCID,cc.OrgName" & vbCrLf
        sSql &= " ,cc.CLASSCNAME2 ClassCName" & vbCrLf
        sSql &= " ,cc.TrainName" & vbCrLf
        sSql &= " ,cc.STDate" & vbCrLf
        sSql &= " ,cc.FTDate" & vbCrLf
        sSql &= " ,cc.RID" & vbCrLf
        sSql &= " ,dbo.NVL(CONVERT(varchar, h.RightID),'XX') RightID" & vbCrLf
        sSql &= " ,dbo.NVL(h.NAME,' ') as NAME,dbo.NVL(h.ACCOUNT,' ') ACCOUNT" & vbCrLf
        sSql &= " ,(CASE WHEN dbo.NVL(h.ACCOUNT,'0') = '0' THEN '0' ELSE '1' END) Acnt" & vbCrLf
        sSql &= " ,h.EndDate" & vbCrLf
        sSql &= " ,h.Temp1" & vbCrLf
        sSql &= " ,OCLASSID ClassID2" & vbCrLf
        sSql &= " FROM dbo.VIEW2 cc" & vbCrLf
        sSql &= " LEFT JOIN (" & vbCrLf
        sSql &= " 	SELECT h1.RightID,h1.OCID,h2.ACCOUNT,h2.NAME,h1.EndDate" & vbCrLf
        sSql &= " 	,d1.Name+';開放FunID: '+h1.FunID Temp1" & vbCrLf
        sSql &= " 	FROM Auth_REndClass h1" & vbCrLf
        sSql &= " 	join Auth_Account h2 ON h1.ACCOUNT = h2.ACCOUNT" & vbCrLf
        sSql &= " 	left join ID_keyinReason d1 on h1.ReasonID=d1.ReasonID" & vbCrLf
        sSql &= " 	WHERE h1.UseAble = 'Y'" & vbCrLf
        sSql &= " ) h ON h.OCID=cc.OCID" & vbCrLf
        sSql &= " WHERE 1=1" & vbCrLf
        sSql &= " " & vbCrLf

        '20090617 andy  edit
        '--------------------
        sSql += "  and cc.IsSuccess='Y'" & vbCrLf '是否轉入成功
        sSql += "  and cc.NotOpen='N' " & vbCrLf  '不開班
        sSql += "  and cc.Years = '" & Me.yearlist.SelectedValue & "'" & vbCrLf
        sSql += "  and cc.DistID = '" & Me.DistID.SelectedValue & "'" & vbCrLf

        If Me.planlist.SelectedValue <> "" Then
            sSql += " and cc.PlanID = '" & Me.planlist.SelectedValue & "' " & vbCrLf
        End If

        dt = DbAccess.GetDataTable(sSql, objconn)
        Return dt
    End Function

    '重新查詢。
    Sub dt_search()
        'TPanel.Visible = False
        'DG_ClassInfo.Visible = False 'TPanel@DG_ClassInfo
        'PageControler1.Visible = False 'TPanel@PageControler1
        'Reason_tr.Visible = False
        'Account_tr.Visible = False
        'Me.trOrgName.Visible = False
        'Label2.Visible = False
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
        MakeddlOrgName(Me.ddlOrgName, Me.planlist.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, Me.DistID.SelectedValue)
        '帳號
        Call MakeAccount(Me.Account, Me.planlist.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, Me.DistID.SelectedValue, Me.ddlOrgName.SelectedValue)
        If vsOrgName <> "" Then
            Common.SetListItem(ddlOrgName, vsOrgName)
        End If
        If vsAccount <> "" Then
            Common.SetListItem(Account, vsAccount)
        End If

        Dim dt As DataTable = Nothing
        dt = ShowDG_ClassInfo(dt)

        'TPanel.Visible = False
        'DG_ClassInfo.Visible = False 'TPanel@DG_ClassInfo
        'PageControler1.Visible = False 'TPanel@PageControler1
        'Reason_tr.Visible = False
        'Account_tr.Visible = False
        'Me.trOrgName.Visible = False
        'Label2.Visible = False
        '------------------20090226 andy add
        Fun_tr.Visible = False
        '----------------
        msg.Text = "查無資料!!"

        If dt.Rows.Count > 0 Then
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()

            'TPanel.Visible = True
            'DG_ClassInfo.Visible = True 'TPanel@DG_ClassInfo
            'PageControler1.Visible = True 'TPanel@PageControler1
            'Reason_tr.Visible = True
            'Account_tr.Visible = True

            'Me.trOrgName.Visible = True
            'Label2.Visible = True
            '------------------20090226 andy add
            Fun_tr.Visible = True
            '----------------
            msg.Text = ""
        End If
    End Sub

    Private Sub DG_ClassInfo_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) 'Handles DG_ClassInfo.ItemCommand

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
        Dim strFDate As String = e.Item.Cells(Cst_FTDate).Text
        Dim strOCID As String = e.Item.Cells(Cst_OCID).Text
        Dim strRightID As String = e.Item.Cells(Cst_RightID).Text

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
                sql &= " INSERT INTO Auth_REndClass ( RightID,Years,PlanID,DistID,OCID,Account,CreateDate,UseAble,ModifyAcct,ModifyDate,Reason,ReasonID,FunID,EndDate)"
                sql &= " VALUES ( @RightID,@Years,@PlanID,@DistID,@OCID,@Account,getdate(),@UseAble,@ModifyAcct, getdate(),@Reason,@ReasonID,@FunID,@EndDate)"
                'sql = ""
                'sql &= " INSERT INTO Auth_REndClass "
                'sql += " (RightID,Years,PlanID,DistID,OCID,Account,CreateDate,UseAble,ModifyAcct,ModifyDate,Reason,ReasonID,FunID,EndDate) "
                'sql += " values(" & Auth_REndClass_MaxNo() & ",'" & Me.yearlist.SelectedValue & "','" & Me.planlist.SelectedValue & "' "
                'sql += " ,'" & Me.DistID.SelectedValue & "','" & strOCID & "','" & Me.Account.SelectedValue & "',getdate(),'Y' "
                'sql += " ,'" & sm.UserInfo.UserID & "',getdate(), @Reason ,'" & Me.ReasonID.SelectedValue & "', '" & chkSelFunID() & "', " & TIMS.to_date(Me.EndDate.Text) & " )"

                Call TIMS.OpenDbConn(objconn)
                Dim oCmd As New SqlCommand(sql, objconn)
                With oCmd
                    .Parameters.Clear()
                    .Parameters.Add("RightID", SqlDbType.Int).Value = Auth_REndClass_MaxNo()
                    .Parameters.Add("Years", SqlDbType.VarChar).Value = Me.yearlist.SelectedValue
                    .Parameters.Add("PlanID", SqlDbType.VarChar).Value = Me.planlist.SelectedValue
                    .Parameters.Add("DistID", SqlDbType.VarChar).Value = Me.DistID.SelectedValue
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = strOCID
                    .Parameters.Add("Account", SqlDbType.VarChar).Value = Me.Account.SelectedValue

                    .Parameters.Add("UseAble", SqlDbType.VarChar).Value = "Y"
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID

                    '.Parameters.Add("Reason", SqlDbType.NVarChar).Value = Me.Reason.Text
                    '.Parameters.Add("ReasonID", SqlDbType.VarChar).Value = Me.ReasonID.SelectedValue
                    '.Parameters.Add("FunID", SqlDbType.VarChar).Value = chkSelFunID()
                    '.Parameters.Add("EndDate", SqlDbType.DateTime).Value = TIMS.cdate2(Me.EndDate.Text)

                    .ExecuteNonQuery()
                End With
                'DbAccess.ExecuteNonQuery(sql, objconn)
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

                sql = ""
                sql &= " UPDATE Auth_REndClass"
                sql += " SET Account =@Account"
                sql += " ,Reason =@Reason"
                sql += " ,ReasonID =@ReasonID"
                sql += " ,ModifyAcct =@ModifyAcct"
                sql += " ,ModifyDate = getdate()"
                sql += " ,FunID = @FunID"
                sql += " ,EndDate= @EndDate"
                sql += " WHERE RightID= @RightID"
                sql += " And OCID= @OCID"
                'sql = ""
                'sql &= " UPDATE Auth_REndClass "
                'sql += " SET Account = '" & Me.Account.SelectedValue & "'"
                'sql += " ,Reason = @Reason "
                'sql += " ,ReasonID = '" & Me.ReasonID.SelectedValue & "'"
                'sql += " ,ModifyAcct = '" & sm.UserInfo.UserID & "'"
                'sql += " ,ModifyDate = getdate()"
                'sql += " ,FunID='" & chkSelFunID() & "'"
                'sql += " ,EndDate=" & TIMS.cdate2(Me.EndDate.Text)
                'sql += " WHERE RightID = '" & strRightID & "' "
                'sql += " And OCID = '" & strOCID & "' "
                Call TIMS.OpenDbConn(objconn)
                Dim oCmd As New SqlCommand(sql, objconn)
                With oCmd
                    .Parameters.Clear()
                    .Parameters.Add("Account", SqlDbType.VarChar).Value = Me.Account.SelectedValue
                    '.Parameters.Add("Reason", SqlDbType.NVarChar).Value = Me.Reason.Text
                    '.Parameters.Add("ReasonID", SqlDbType.VarChar).Value = Me.ReasonID.SelectedValue
                    '.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    '.Parameters.Add("FunID", SqlDbType.VarChar).Value = chkSelFunID()
                    '.Parameters.Add("EndDate", SqlDbType.DateTime).Value = TIMS.cdate2(Me.EndDate.Text)
                    .Parameters.Add("RightID", SqlDbType.VarChar).Value = strRightID
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = strOCID
                    .ExecuteNonQuery()
                End With

                Common.MessageBox(Me, "修改成功")

                Call dt_search()
            Case "Del"

                If strRightID = "XX" Then
                    Common.MessageBox(Me, "目前尚未針對本班級授權,無法進行刪除動作！")
                    Exit Sub
                End If
                sql = ""
                sql &= " UPDATE Auth_REndClass "
                sql += " SET UseAble = 'N'"
                sql += " ,ModifyAcct = '" & sm.UserInfo.UserID & "'"
                sql += " ,ModifyDate = getdate() "
                sql += " WHERE RightID = '" & TIMS.ClearSQM(strRightID) & "' "
                sql += " And OCID = '" & TIMS.ClearSQM(strOCID) & "' "
                DbAccess.ExecuteNonQuery(sql, objconn)
                Common.MessageBox(Me, "刪除成功")

                Call dt_search()

            Case "GetData"
                If strRightID = "XX" Then
                    Common.MessageBox(Me, "目前尚未針對本班級授權,無法進行取得動作！")
                    Exit Sub
                End If
                sql = "select * from Auth_REndClass WHERE RightID = '" & strRightID & "' "
                Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
                If dt.Rows.Count > 0 Then
                    Dim dr As DataRow = dt.Rows(0)
                    Common.SetListItem(yearlist, dr("Years"))
                    Common.SetListItem(planlist, dr("PlanID"))
                    Common.SetListItem(DistID, dr("DistID"))

                    Common.SetListItem(Account, dr("Account"))
                    'Common.SetListItem(ReasonID, dr("ReasonID"))
                    'Reason.Text = Convert.ToString(dr("Reason"))
                    'Me.EndDate.Text = TIMS.cdate3(dr("EndDate"))
                    If Convert.ToString(dr("FunID")) <> "" Then
                        For i As Int16 = 0 To cb_SelFunID.Items.Count - 1
                            cb_SelFunID.Items(i).Selected = False
                            If Convert.ToString(dr("FunID")).IndexOf(cb_SelFunID.Items(i).Value) > -1 Then
                                cb_SelFunID.Items(i).Selected = True
                            End If
                        Next
                    End If
                End If

        End Select

    End Sub

    Private Sub DG_ClassInfo_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) 'Handles DG_ClassInfo.ItemDataBound
        Dim drv As DataRowView = e.Item.DataItem

        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim but1 As LinkButton = e.Item.FindControl("but1") '新增
                Dim but2 As LinkButton = e.Item.FindControl("but2") '修改
                Dim but3 As LinkButton = e.Item.FindControl("but3") '刪除
                Dim but4 As LinkButton = e.Item.FindControl("but4") '取得

                'Dim but4 As Button = e.Item.FindControl("but4") '查看
                'but4.CommandArgument = Convert.ToString(drv("RightID"))
                If Convert.ToString(drv("Temp1")) <> "" Then
                    TIMS.Tooltip(e.Item, Convert.ToString(drv("Temp1")))
                    TIMS.Tooltip(but1, Convert.ToString(drv("Temp1")))
                    TIMS.Tooltip(but2, Convert.ToString(drv("Temp1")))
                    TIMS.Tooltip(but3, Convert.ToString(drv("Temp1")))
                End If

                but3.Attributes("onclick") = "return confirm('確定要刪除此授權?');"
                'but3.Attributes("onclick") = "return ChkData();return confirm('確定要刪除此授權?');"   '20090226 andy edit

                '20090226 andy edit
                '---------------------
                'but1.Attributes("onclick") = "return ChkData();"
                'but2.Attributes("onclick") = "return ChkData();"
                '---------------------
        End Select
    End Sub

    '查詢
    Private Sub rt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rt_search.Click
        dt_search()
    End Sub

    '依計畫 查詢
    Private Sub planlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles planlist.SelectedIndexChanged
        dt_search()
    End Sub

    Private Sub ddlOrgName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddlOrgName.SelectedIndexChanged
        '帳號
        Call MakeAccount(Me.Account, Me.planlist.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, Me.DistID.SelectedValue, Me.ddlOrgName.SelectedValue)

        'Dim dt As DataTable = Nothing
        'dt = ShowDG_ClassInfo(dt)
        'DG_ClassInfo.Visible = False 'TPanel@DG_ClassInfo
        'PageControler1.Visible = False 'TPanel@PageControler1
        'msg.Text = "查無資料!!"
        'If dt.Rows.Count > 0 Then
        '    PageControler1.PageDataTable = dt
        '    PageControler1.ControlerLoad()
        '    DG_ClassInfo.Visible = True 'TPanel@DG_ClassInfo
        '    PageControler1.Visible = True 'TPanel@PageControler1
        '    msg.Text = ""
        'End If

    End Sub

End Class