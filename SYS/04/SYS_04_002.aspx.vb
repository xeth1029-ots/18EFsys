Partial Class SYS_04_002
    Inherits AuthBasePage

#Region "(目前不使用)"

    'Dim objreader As SqlDataReader
    'Dim FunDr As DataRow
    '配合TABLE : KEY_TABLEMGR (KEYTABLE:大小寫有差程式有寫死的判斷)
    'EXAMPLE:
    'INSERT INTO KEY_TABLEMGR(SERNUM ,KEYTYPE,KEYTABLE) VALUES (47,'選擇郵遞區號','ID_ZIP2')
    'select max(SERNUM )+1 mmmSERNUM FROM KEY_TABLEMGR
    'SELECT SERNUM ,KEYTYPE,KEYTABLE FROM KEY_TABLEMGR WHERE SERNUM>=47-10 ORDER BY SERNUM 
    'Dim sqlstr As String

#End Region

    Const Cst_defaultQID As String = "3"

    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        Call TB_DISPLAY_NONE()

        'RegularExpressionValidator1.Enabled = True

        If Not IsPostBack Then
            Call Create1()
        End If

        tvUnit.Nodes.Clear()

        Call AddJavaScript()

#Region "目前不使用"

        ''檢查帳號的功能權限-----------------------------------Start
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
        '            If FunDr("Adds") = "1" Then
        '                cmdAppend.Enabled = True
        '            Else
        '                cmdAppend.Enabled = False
        '            End If
        '            If FunDr("Sech") = "1" Then
        '                btnSearch.Enabled = True
        '            Else
        '                btnSearch.Enabled = False
        '            End If
        '            If FunDr("Mod") = "1" Then
        '                cmdUpdate.Enabled = True
        '            Else
        '                cmdUpdate.Enabled = False
        '            End If
        '        End If
        '    End If
        'End If
        'cmdAppend.Enabled = True
        'If Not blnCanAdds Then cmdAppend.Enabled = False
        'btnSearch.Enabled = True
        'If Not blnCanSech Then btnSearch.Enabled = False
        'cmdUpdate.Enabled = True
        'If Not blnCanMod Then cmdUpdate.Enabled = False
#End Region
        ViewState("sPath") = TIMS.Server_Path
        '檢查帳號的功能權限 End
    End Sub

    ''' <summary>初始時，隱藏所有TABLE</summary>
    Private Sub TB_DISPLAY_NONE()
        btnConfigReset.Visible = False
        Table4.Visible = False
        Table5.Visible = False
        Table2.Style.Item("display") = "none"
        Table6.Style.Item("display") = "none"
        Table8.Style.Item("display") = "none"
        Table7.Style.Item("display") = "none"
        Table9.Style.Item("display") = "none"
        Table10.Style.Item("display") = "none"
        Table11.Style.Item("display") = "none"
        Table11b.Style.Item("display") = "none"
        Table12.Style.Item("display") = "none"
        Table13.Style.Item("display") = "none" 'Key_Identity
        Table14.Style.Item("display") = "none"
        Table15.Style.Item("display") = "none"
        Table16.Style.Item("display") = "none"
        Table17.Style.Item("display") = "none"

        'keyUnit.Attributes.Add("style", "BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-WEIGHT: normal; FONT-SIZE: 9px; OVERFLOW: scroll; BORDER-LEFT: #000000 1px solid; WIDTH: 250px; COLOR: #fcefc7; BORDER-BOTTOM: #000000 1px solid; FONT-FAMILY: 新細明體; HEIGHT: 600px; BACKGROUND-COLOR: #ffffcc")
        keyUnit.Attributes.Add("style", " BORDER-LEFT: #000000 1px solid; BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid; OVERFLOW: scroll; FONT-FAMILY: 新細明體; HEIGHT: 600px;  WIDTH: 300px; COLOR: #fcefc7; BACKGROUND-COLOR: #ffffcc")
    End Sub

    Sub Create1()
        'Dim sqlstr As String
        'sqlstr = "select KeyTable,KeyType from Key_TableMgr ORDER BY KeyTable "
        Dim sqlstr As String = "SELECT KEYTABLE,KEYTYPE,SERNUM FROM KEY_TABLEMGR ORDER BY SERNUM,KEYTABLE"
        DbAccess.MakeListItem(KeyType, sqlstr, objconn)

        sqlstr = "SELECT * FROM KEY_TRAINTYPE WHERE LEVELS=0 ORDER BY BUSID"
        Dim dt As DataTable = DbAccess.GetDataTable(sqlstr, objconn)
        With BusID
            .DataSource = dt
            .DataTextField = "BusName"
            .DataValueField = "BusID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With

        '跨計畫合併限制處分紀錄，因跨計畫合併限制處分紀錄可能會有不同組合，需要另外一個欄位紀錄組合喔
        cblSBK2TPLANID = TIMS.Get_TPlan(cblSBK2TPLANID, TIMS.dtNothing(), 0, "N", "", objconn)

        KeyType.Attributes("onchange") = "display_Item();"
        BusID.Attributes("onchange") = "Train(0);"
        JobID.Attributes("onchange") = "Get_TMID();"
        PlanType.Attributes("onchange") = "document.FDUpdate.TypeValue.value=this.value;"

        BusID.Style.Item("display") = "none"
        JobID.Style.Item("display") = "none"

        btnSearch.Attributes("onclick") = "return check_search()"

        rblEmailSend.Attributes("onclick") = "ChangerblReusable();"
        'ddlUnUsedYear.Attributes("onclick") = "ChangerblReusable();"
        ddlUnUsedYear.Attributes("onchange") = "ChangeddlMergeID();"
        cmdUpdate.Attributes("onclick") = "return check_save();"

        '選擇全部轄區
        IdentityID.Attributes("onclick") = "SelectAll('IdentityID','IdentityIDHidden');"
    End Sub

    Sub AddJavaScript()
        'Dim dr3() As DataRow
        'Dim i As Integer, j As Integer, k As Integer
        Dim dt As DataTable
        Dim sqlstr As String = "SELECT * FROM KEY_TRAINTYPE ORDER BY TMID"
        dt = DbAccess.GetDataTable(sqlstr, objconn)
        Dim dr1() As DataRow
        dr1 = dt.Select("Levels=0")

        Dim JScript As String = ""
        JScript += "<script language='javascript'>" & vbCrLf
        JScript &= "    function Train(num){" & vbCrLf
        For i As Integer = 0 To dr1.Length - 1
            Dim myName As String = ""
            Dim myValue As String = ""
            'dr2 = dt.Select("Parent=1")
            Dim dr2 As DataRow() = dt.Select("Levels=1 and [Parent]='" & dr1(i)("TMID") & "'")
            For j As Integer = 0 To dr2.Length - 1
                myName &= String.Concat(If(myName <> "", ",", ""), "'", dr2(j)("JobName"), "'")
                myValue &= String.Concat(If(myValue <> "", ",", ""), "'", dr2(j)("JobID"), "'")
            Next
            JScript += String.Concat(" var ", dr1(i)("BusID"), "1= new Array(", myName, ");") & vbCrLf
            JScript += String.Concat(" var ", dr1(i)("BusID"), "2= new Array(", myValue, ");") & vbCrLf
        Next

        JScript &= "        var mydrop=document.getElementById('KeyType');" & vbCrLf
        JScript &= "        if(mydrop.value=='Key_TrainType3'){ //表示要開始做處裡" & vbCrLf
        JScript &= "            var mydrop=document.getElementById('BusID');" & vbCrLf
        JScript &= "            if(mydrop.value!=''){" & vbCrLf
        JScript &= "                //先清除所有的值" & vbCrLf
        JScript &= "                for(var i=document.FDUpdate.JobID.options.length;i>0;i--){" & vbCrLf
        JScript &= "                    document.FDUpdate.JobID.options[i]=null;" & vbCrLf
        JScript &= "                    document.FDUpdate.TMIDValue.value='';" & vbCrLf
        JScript &= "                }" & vbCrLf
        JScript &= "                //開始輸入新值" & vbCrLf
        JScript &= "                document.FDUpdate.JobID.options[0]=new Option('===請選擇===','');" & vbCrLf
        JScript &= "                //判斷父節點" & vbCrLf
        JScript &= "                var parentName=eval(mydrop.value+'1');" & vbCrLf
        JScript &= "                var parentValue=eval(mydrop.value+'2');" & vbCrLf
        JScript &= "                for (var i=0;i<parentName.length;i++){" & vbCrLf
        JScript &= "                    document.FDUpdate.JobID.options[i+1]=new Option(parentName[i],parentValue[i]);" & vbCrLf
        JScript &= "                }" & vbCrLf
        JScript &= "                //如果之前有輸入，保留欄位" & vbCrLf
        JScript &= "                if(document.FDUpdate.TMIDValue.value!=''){" & vbCrLf
        JScript &= "                    var mydrop1=document.getElementById('JobID');" & vbCrLf
        JScript &= "                    mydrop1.value=document.FDUpdate.TMIDValue.value;" & vbCrLf
        JScript &= "                }" & vbCrLf
        JScript &= "            }" & vbCrLf
        JScript &= "        }" & vbCrLf
        JScript &= "    }" & vbCrLf
        JScript += "</script>" & vbCrLf

        Common.RespWrite(Me, JScript)
    End Sub

    'Update Plan_Identity Key_Plan
    Public Shared Sub UPDATE_PLAN_IDENTITY(ByRef sm As SessionModel, ByRef IdentityID As CheckBoxList, ByVal TPlanID As String, ByRef oConn As SqlConnection)
        'TIMS.OpenDbConn(oConn)
        'Dim sql As String
        Dim dt2 As New DataTable '= Nothing
        'Dim da2 As SqlDataAdapter = Nothing
        Dim sqlstr As String = ""
        sqlstr = " SELECT * FROM PLAN_IDENTITY WHERE TPlanID=@TPlanID" '" & TPlanID & "'"
        Dim sCmd As New SqlCommand(sqlstr, oConn)

        Dim i_sql As String = ""
        i_sql &= " INSERT INTO PLAN_IDENTITY(TPLANID ,IDENTITYID ,ISENABLED ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
        i_sql &= " VALUES (@TPlanID ,@IdentityID ,@ISENABLED ,@MODIFYACCT ,getdate())" & vbCrLf
        Dim iCmd As New SqlCommand(i_sql, oConn)

        Dim u_sql As String = ""
        u_sql &= " UPDATE PLAN_IDENTITY" & vbCrLf
        u_sql &= " SET ISENABLED=@ISENABLED ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=getdate()" & vbCrLf
        u_sql &= " WHERE 1=1 AND TPLANID=@TPlanID AND IDENTITYID=@IdentityID" & vbCrLf
        Dim uCmd As New SqlCommand(u_sql, oConn)

        With sCmd
            .Parameters.Clear()
            .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = TPlanID
            dt2.Load(.ExecuteReader())
        End With
        For i As Integer = 1 To IdentityID.Items.Count - 1
            'Dim dr2 As DataRow = Nothing
            Dim v_ISENABLED As String = If(IdentityID.Items(i).Selected, "Y", "N")
            If dt2.Select("IdentityID='" & IdentityID.Items(i).Value & "'").Length > 0 Then
                '修改
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("ISENABLED", SqlDbType.VarChar).Value = v_ISENABLED
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = TPlanID
                    .Parameters.Add("IdentityID", SqlDbType.VarChar).Value = IdentityID.Items(i).Value
                    .ExecuteNonQuery()
                End With
            Else
                '新增修改
                With iCmd
                    .Parameters.Clear()
                    .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = TPlanID
                    .Parameters.Add("IdentityID", SqlDbType.VarChar).Value = IdentityID.Items(i).Value
                    .Parameters.Add("ISENABLED", SqlDbType.VarChar).Value = v_ISENABLED
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .ExecuteNonQuery()
                End With
            End If

        Next
        'DbAccess.UpdateDataTable(dt2, da2)
    End Sub

    Function SUtl_GetMsg1(ByVal tableName As String) As String
        Return String.Concat("因 ", tableName, " Table結構改變，故不可使用此相同規則修改，請洽管理者協助修改")
    End Function

    ''' <summary>清理物件值</summary>
    Sub ClearObjValue()
        keycode.Text = ""
        keyname.Text = ""
        txtmemo1.Text = ""
        Levels.Value = ""
        Parent1.Value = ""
        MinusPoint.Text = ""
        cb_LEAVE_NOUSE.Checked = False
        'Sort.Text = ""
        EngkeyName.Text = ""
        AddMinus.Text = ""
        point.Text = ""
        Sort.Text = ""
        DGHour.Text = ""
        ClsYear.Text = ""
        txtItemageName.Text = ""
        txtItemCostName.Text = ""
        txtKeyTable.Text = ""
        MONTHLYWAGE.Text = ""
        HOURLYWAGE.Text = ""
        SWDATE.Text = ""
        hid_BASICSALARY_SBSID.Value = ""
        lab_BASICSALARY_SBSID.Text = ""

        'Table17
        CB_MUSTFILL.Checked = False
        cb_USELATESTVER.Checked = False
        cb_DOWNLOADRPT.Checked = False
        cb_UPLOADFL1.Checked = False
        cb_SENTBATVER.Checked = False
        cb_USEMEMO1.Checked = False
        'cb_DataGrid08.Checked = False
        If RBL_ORGKINDGW.SelectedItem IsNot Nothing Then RBL_ORGKINDGW.SelectedItem.Selected = False
        txt_KSORT.Text = ""
        txt_RPTNAME.Text = ""
        txt_KBDESC1.Text = ""

        If IsOnLine.SelectedItem IsNot Nothing Then IsOnLine.SelectedItem.Selected = False
        If PlanType.SelectedItem Is Nothing Then PlanType.SelectedItem.Selected = False
    End Sub

    ''' <summary>查詢</summary>
    Sub SbSearch1()
        Call ClearObjValue()

        Dim V_sPath As String = Convert.ToString(ViewState("sPath")) 'Convert.ToString(dr("QID"))

        Dim strSql As String = ""
        '取得值
        Dim v_KeyType As String = TIMS.GetListValue(KeyType)
        Select Case v_KeyType 'v_KeyType 'KeyType.SelectedValue
            Case "KEY_BIDCASE"
                Dim sSql As String = ""
                sSql &= " SELECT KBSID,KBID,KBNAME,KBDESC1,MUSTFILL,ORGKINDGW" & vbCrLf
                sSql &= " ,KSORT,USELATESTVER,DOWNLOADRPT,RPTNAME,UPLOADFL1,SENTBATVER,USEMEMO1" & vbCrLf
                sSql &= " FROM KEY_BIDCASE" & vbCrLf
                sSql &= " ORDER BY KSORT,KBID,KBSID" & vbCrLf
                strSql = sSql
                Table3.Style.Item("display") = "" '鍵值代碼/鍵值名稱
                Table17.Style.Item("display") = ""

            Case "SYS_BASICSALARY"
                'https://www.mol.gov.tw/1607/28162/28166/28180/28182/
                Table3.Style.Item("display") = "none" '鍵值代碼/鍵值名稱
                Table16.Style.Item("display") = ""
                strSql = "SELECT SBSID,MONTHLYWAGE,HOURLYWAGE,FORMAT(SWDATE,'yyyy/MM/dd') SWDATE FROM SYS_BASICSALARY ORDER BY SBSID DESC"
            Case "SYS_VAR"
                btnConfigReset.Visible = True
                'RegularExpressionValidator1.Enabled = False
                Table15.Style.Item("display") = ""
                strSql = "SELECT SVID,ITEMNAME KID,ITEMVALUE NAME,MEMO1 FROM SYS_VAR WHERE SPAGE='Config'"
                strSql &= " ORDER BY 1" & vbCrLf
            Case "ID_DISTRICT"
                strSql = "SELECT DISTID KID,Name From ID_DISTRICT"
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_OrgType"
                strSql = "SELECT OrgTypeID AS KID,Name From Key_OrgType"
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_Degree"
                'Table7.Style.Item("display") = "inline" '附加
                'Table10.Style.Item("display") = "inline" '附加
                Table7.Style.Item("display") = ""
                Table10.Style.Item("display") = ""
                strSql = "SELECT DegreeID AS KID,Name, DegreeType, Sort From Key_Degree"
                strSql &= " ORDER BY Sort,DegreeID" & vbCrLf
            Case "Key_Identity"
                strSql = ""
                strSql &= " SELECT IdentityID AS KID,Name,UnUsedYear,MergeID,Subsidy"
                strSql &= " ,SORT28,SUPPLYID,NOSHOWMI"
                strSql &= " FROM Key_Identity"
                strSql &= " ORDER BY 1" & vbCrLf
                'Table13.Style.Item("display") = "inline" '附加
                Table13.Style.Item("display") = ""
                '停用年度
                If ddlUnUsedYear.SelectedIndex = -1 Then
                    ddlUnUsedYear = TIMS.Get_Years(ddlUnUsedYear, objconn)
                    ddlUnUsedYear.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                End If
                '併入身分別
                If ddlMergeID.SelectedIndex = -1 Then
                    ddlMergeID = TIMS.Get_Identity(ddlMergeID, 2, objconn)
                End If

            Case "Key_Military"
                strSql = "SELECT MilitaryID AS KID,Name From Key_Military"
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_Plan"
                'Table8.Style.Item("display") = "inline" '附加
                'Table14.Style.Item("display") = "inline" '附加
                Table8.Style.Item("display") = ""
                Table14.Style.Item("display") = ""

                '取出鍵詞-身分別代碼 Type:3
                IdentityID = TIMS.Get_Identity(IdentityID, 3, objconn)

                ' strSql = "Select TPlanID AS KID,PlanName AS Name,IsOnLine,PlanType,ClsYear, QuesID From Key_Plan"
                ' TIMS.Get_KeyQuesType(QuesType)
                'Dim sql As String = ""
                strSql = ""
                strSql &= " SELECT a.TPlanID KID, a.PlanName NAME" & vbCrLf
                strSql &= " ,a.IsOnLine,a.PlanType,a.ClsYear" & vbCrLf
                strSql &= " ,a.EmailSend,a.Reusable,a.BlackList" & vbCrLf
                strSql &= " ,a.QueryDisplay" & vbCrLf
                strSql &= " ,ISNULL(a.useECFA,'N') useECFA" & vbCrLf
                strSql &= " ,a.PropertyID" & vbCrLf
                strSql &= " ,b.QID,c.QName" & vbCrLf 'TIMS
                strSql &= " ,a.QuesID,u.Name QNameU" & vbCrLf 'DEMO
                strSql &= " ,'0'+dbo.FN_GET_IDENTITYID(a.TPlanID) IdentityID" & vbCrLf
                strSql &= " ,ISNULL(a.SBLACKTYPE,0) SBLACKTYPE" & vbCrLf
                strSql &= " ,dbo.FN_GET_SBLACK2TPLANID(a.TPlanID) SBK2TPLANID" & vbCrLf
                'strSql &= " ,a.SBLACK2TPLANID" & vbCrLf
                strSql &= " FROM KEY_PLAN a" & vbCrLf
                strSql &= " LEFT JOIN Plan_Questionary b ON a.TPlanID = b.TPlanID" & vbCrLf
                strSql &= " LEFT JOIN ID_Questionary c ON b.QID=c.QID" & vbCrLf
                strSql &= " LEFT JOIN ID_Survey u ON u.SVID=a.QuesID" & vbCrLf
                strSql &= " ORDER BY a.TPlanID" & vbCrLf

                Select Case ViewState("sPath")
                    Case "TIMS"
                        '20080729 andy
                        'strSql = "" & vbCrLf
                        'strSql &= " SELECT a.TPlanID AS KID, a.PlanName AS Name" & vbCrLf
                        'strSql &= " ,a.IsOnLine,a.PlanType,a.ClsYear" & vbCrLf
                        'strSql &= " ,a.EmailSend,a.Reusable,a.BlackList,a.QueryDisplay" & vbCrLf
                        'strSql &= " ,ISNULL(a.useECFA,'N') useECFA" & vbCrLf
                        'strSql &= " ,a.PropertyID" & vbCrLf
                        'strSql &= " ,b.QID, c.QName" & vbCrLf
                        'strSql &= " ,'0'+dbo.FN_GET_IDENTITYID(a.TPlanID) IdentityID" & vbCrLf
                        'strSql &= " FROM Key_Plan a" & vbCrLf
                        'strSql &= " LEFT JOIN Plan_Questionary b ON a.TPlanID = b.TPlanID" & vbCrLf
                        'strSql &= " LEFT JOIN ID_Questionary c ON b.QID=c.QID" & vbCrLf
                        'strSql &= " ORDER BY a.TPlanID" & vbCrLf
                        '放入目前計畫問卷別
                        TIMS.Get_ID_Questionary(QuesType)
                        '預設QID  'Dim defalutQID As Integer = "3"
                        Common.SetListItem(QuesType, Cst_defaultQID)

                    Case "DEMO"
                        '採用自動問卷功能
                        Dim Survey_sql As String = ""
                        Survey_sql = "SELECT SVID AS QID,Name AS QName FROM id_Survey WHERE Avail='Y'"
                        Survey_sql &= " ORDER BY 1" & vbCrLf
                        TIMS.Get_ID_Questionary(QuesType, DbAccess.GetDataTable(Survey_sql, objconn)) '放入目前計畫問卷別
                        '=====
                        'strSql = "" & vbCrLf
                        'strSql &= " SELECT a.TPlanID AS KID, a.PlanName AS Name" & vbCrLf
                        'strSql &= " ,a.IsOnLine,a.PlanType,a.ClsYear " & vbCrLf
                        'strSql &= " ,a.EmailSend,a.Reusable,a.BlackList ,a.QueryDisplay " & vbCrLf
                        'strSql &= " ,ISNULL(a.useECFA,'N') useECFA" & vbCrLf
                        'strSql &= " ,a.PropertyID" & vbCrLf
                        'strSql &= " ,a.QuesID, c.Name AS QName  " & vbCrLf
                        'strSql &= " ,'0'+dbo.FN_GET_IDENTITYID( a.TPlanID ) IdentityID" & vbCrLf
                        'strSql &= " FROM Key_Plan a " & vbCrLf
                        'strSql &= " LEFT JOIN ID_Survey c ON c.SVID=a.QuesID" & vbCrLf
                        'strSql &= " ORDER BY 1" & vbCrLf

                    Case Else 'TIMS '其他情況暫用正式機
                        '20080729 andy
                        'strSql = "" & vbCrLf
                        'strSql &= " SELECT a.TPlanID AS KID, a.PlanName AS Name" & vbCrLf
                        'strSql &= " ,a.IsOnLine,a.PlanType,a.ClsYear " & vbCrLf
                        'strSql &= " ,a.EmailSend,a.Reusable,a.BlackList ,a.QueryDisplay " & vbCrLf
                        'strSql &= " ,ISNULL(a.useECFA,'N') useECFA" & vbCrLf
                        'strSql &= " ,a.PropertyID" & vbCrLf
                        'strSql &= " ,b.QID, c.QName " & vbCrLf
                        'strSql &= " ,'0'+dbo.FN_GET_IDENTITYID( a.TPlanID ) IdentityID" & vbCrLf
                        'strSql &= " FROM Key_Plan a " & vbCrLf
                        'strSql &= " LEFT JOIN Plan_Questionary b ON a.TPlanID = b.TPlanID" & vbCrLf
                        'strSql &= " LEFT JOIN ID_Questionary c ON b.QID = c.QID" & vbCrLf
                        'strSql &= " ORDER BY 1" & vbCrLf
                        TIMS.Get_ID_Questionary(QuesType)   '放入目前計畫問卷別
                        '預設QID
                        'Dim defalutQID As Integer = "3"
                        Common.SetListItem(QuesType, Cst_defaultQID)

                End Select
            Case "Key_RejectTReason"
                strSql = "SELECT RTReasonID AS KID,Reason AS Name FROM Key_RejectTReason"
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_SelResult"
                strSql = "SELECT SelResultID AS KID ,Name FROM Key_SelResult"
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_Subsidy"
                strSql = "SELECT SubsidyID AS KID,Name FROM Key_Subsidy"
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_TrainingSource"
                strSql = "SELECT SourceID AS KID,Name FROM Key_TrainingSource"
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_TrainType1"
                'RegularExpressionValidator1.Enabled = False
                strSql = "SELECT BusID AS KID,BusName AS Name FROM Key_TrainType WHERE BusID IS NOT NULL"
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_TrainType2"
                Dim v_BusID As String = TIMS.GetListValue(BusID) '.SelectedValue
                If v_BusID = "" Then
                    Page.RegisterStartupScript("window_onload", "<script language='javascript'>display_Item();</script>")
                    Exit Sub
                End If
                'RegularExpressionValidator1.Enabled = False
                strSql = "SELECT TMID FROM Key_TrainType WHERE BusID = '" & v_BusID & "' "
                Dim dr As DataRow = DbAccess.GetOneRow(strSql, objconn)
                If Not dr Is Nothing Then
                    strSql = "SELECT JobID AS KID,JobName AS Name,TMID,Parent FROM Key_TrainType WHERE Parent='" & dr("TMID") & "' "
                    strSql &= " ORDER BY TMID" & vbCrLf
                    Parent1.Value = dr("TMID")
                    Table4.Visible = True
                    Bus.Text = BusID.SelectedItem.Text
                End If
                Page.RegisterStartupScript("window_onload", "<script language='javascript'>display_Item();</script>")
            Case "Key_TrainType3"
                Dim v_JobID As String = TIMS.GetListValue(JobID) '.SelectedValue
                Dim v_BusID As String = TIMS.GetListValue(BusID) '.SelectedValue
                If v_BusID = "" OrElse v_JobID = "" Then
                    Page.RegisterStartupScript("window_onload", "<script language='javascript'>display_Item();</script>")
                    Exit Sub
                End If
                'RegularExpressionValidator1.Enabled = False
                strSql = "SELECT TMID FROM Key_TrainType WHERE BusID = '" & v_BusID & "' "
                Dim dr As DataRow = DbAccess.GetOneRow(strSql, objconn)
                Dim par As String = dr("TMID")

                strSql = "SELECT JobName,TMID FROM Key_TrainType WHERE JobID = '" & TMIDValue.Value & "' and Parent = '" & dr("TMID") & "' "
                dr = DbAccess.GetOneRow(strSql, objconn)
                If Not dr Is Nothing Then
                    strSql = "SELECT TrainID AS KID,TrainName AS Name,TMID,Parent FROM Key_TrainType WHERE Parent = '" & dr("TMID") & "' "
                    strSql &= " ORDER BY TMID" & vbCrLf
                    Parent1.Value = dr("TMID")
                    Table4.Visible = True
                    Table5.Visible = True
                    Bus.Text = BusID.SelectedItem.Text
                    Job.Text = dr("JobName")
                End If
                Page.RegisterStartupScript("window_onload", "<script language='javascript'>display_Item();</script>")
            Case "Key_Leave"
                strSql = "SELECT LEAVEID AS KID,NAME,MINUSPOINT,NOUSE,LEAVESORT,ENGNAME FROM KEY_LEAVE ORDER BY 1"
                'Table2.Style.Item("display") = "inline" '附加
                Table2.Style.Item("display") = ""
                Table7.Style.Item("display") = ""'sort
            Case "Key_Sanction"
                strSql = "SELECT SanID AS KID,Name,AddMinus,Point FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
                'Table6.Style.Item("display") = "inline" '附加
                Table6.Style.Item("display") = ""
            Case "Key_GradState"
                strSql = "SELECT GradID AS KID,Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_JoblessWeek"
                strSql = "SELECT JoblessID AS KID,Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_HourRan"
                strSql = "SELECT HRID AS KID,HourRanName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_TrainExp"
                strSql = "SELECT TEID AS KID,TrainExpName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_HandicatType"
                strSql = "SELECT HandTypeID AS KID,Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_HandicatLevel"
                strSql = "SELECT HandLevelID AS KID,Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_HandicatType2"
                strSql = "SELECT HandTypeID2 AS KID,Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_HandicatLevel2"
                strSql = "SELECT HandLevelID2 AS KID,Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY HandLev1elID2" & vbCrLf
            Case "Key_Exam"
                strSql = "SELECT ExamID AS KID,Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_CostItem"
                'Table7.Style.Item("display") = "inline" '附加
                'Table11.Style.Item("display") = "inline" '附加
                'Table11b.Style.Item("display") = "inline" '附加
                Table7.Style.Item("display") = ""
                Table11.Style.Item("display") = ""
                Table11b.Style.Item("display") = ""
                strSql = "SELECT CostID AS KID, CostName AS Name, ItemageName, ItemCostName, Sort FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_CostItem2"
                'Table7.Style.Item("display") = "inline" '附加
                'Table11b.Style.Item("display") = "inline" '附加
                Table7.Style.Item("display") = ""
                Table11b.Style.Item("display") = ""
                strSql = "SELECT CostID AS KID, CostName AS Name, ItemCostName, Sort FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_Budget"
                strSql = "SELECT BudID AS KID,BudName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_DGTHour"
                'Table9.Style.Item("display") = "inline" '附加
                Table9.Style.Item("display") = ""
                strSql = "SELECT DGID AS KID,DGName AS Name,DGHour FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_Salary"
                strSql = "SELECT SalID AS KID,SalName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_GetJob"
                strSql = "SELECT GJID AS KID,GJName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_NonGetJob"
                strSql = "SELECT NGJID AS KID,NGJName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_BusVisitCase"
                strSql = "SELECT BVCID AS KID,BVCName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_AgeRange"
                strSql = "SELECT ARID AS KID,ARName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_WorkYear"
                strSql = "SELECT WYID AS KID,WYName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_Emp"
                strSql = "SELECT KEID AS KID,KEName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_KindofJob"
                'strSql = "Select TradeID AS KID,TradeName AS Name From " & v_KeyType 'KeyType.SelectedValue
                Dim msg As String = SUtl_GetMsg1(v_KeyType) 'KeyType.SelectedValue)
                Common.MessageBox(Me, msg)
                Exit Sub
            Case "Key_ProSkill"
                'strSql = "Select TradeID AS KID,TradeName AS Name From " & v_KeyType 'KeyType.SelectedValue
                Dim msg As String = SUtl_GetMsg1(v_KeyType) 'KeyType.SelectedValue)
                Common.MessageBox(Me, msg)
                Exit Sub
            Case "Key_Trade"
                strSql = "SELECT TradeID AS KID,TradeName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_NotOpenReason"
                strSql = "SELECT NORID AS KID,NORName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_ClassCatelog"
                strSql = "SELECT CCID AS KID,CCName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_CheckRan"
                strSql = "SELECT CHID AS KID,CheckRanName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "ID_Invest"
                strSql = "SELECT IVID AS KID,InvestName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "OB_Funds"
                strSql = "SELECT FID AS KID, FName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_Native"
                strSql = "SELECT KNID AS KID, Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_SurveyType"
                strSql = "SELECT STID AS KID, STName AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_JobGroup" '職群鍵詞檔
                'Table7.Style.Item("display") = "inline" '附加
                Table7.Style.Item("display") = ""
                strSql = "SELECT JGID KID, JGNAME AS Name ,Sort FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_EnterPoint" '錄訓百分比代碼
                strSql = "SELECT KID, KNAME AS Name FROM " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
            Case "Key_TableMgr"
                'Table12.Style.Item("display") = "inline" '附加
                Table12.Style.Item("display") = ""
                strSql = "SELECT SerNum AS KID,KeyType AS Name , KeyTable From " & v_KeyType 'KeyType.SelectedValue
                strSql &= " ORDER BY 1" & vbCrLf
        End Select

        Dim objAdapter As SqlDataAdapter = Nothing
        Dim objTable As DataTable = Nothing
        If strSql = "" Then
            '沒有建立可用的SQL語法
            Common.MessageBox(Me, "此鍵詞功能尚未完成，請連絡系統人員建立...")
            Exit Sub
        End If
        objTable = DbAccess.GetDataTable(strSql, objAdapter, objconn)
        For Each row As DataRow In objTable.Rows
            AddTreeNodes(row, objTable, tvUnit, Nothing)
        Next

    End Sub

    '查詢
    Private Sub BtnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        SbSearch1()
    End Sub

    '加入 TreeNodes
    Sub AddTreeNodes(ByVal dr As DataRow, ByVal objTable As DataTable, ByVal objTreeView As TreeView, ByVal ParentNode As TreeNode)
        Dim NewNode As New TreeNode
        'Dim drChild As DataRow
        'Dim strFilter As String
        Dim s_TMP_URL_1 As String = ""
        Dim V_sPath As String = Convert.ToString(ViewState("sPath")) 'Convert.ToString(dr("QID"))
        Dim v_KeyType As String = TIMS.GetListValue(KeyType) 'KeyType.SelectedValue
        Select Case v_KeyType 'KeyType.SelectedValue
            Case "KEY_BIDCASE"
                NewNode.Text = String.Concat("(", dr("KBSID"), ")", dr("ORGKINDGW"), dr("KBID"), ".", dr("KBNAME"))
            Case "SYS_BASICSALARY"
                NewNode.Text = String.Concat("(", dr("SBSID"), ")", TIMS.Cdate3(dr("SWDATE")))
            Case Else
                NewNode.Text = String.Concat("(", dr("KID"), ")", dr("NAME"))
        End Select

        Select Case v_KeyType 'KeyType.SelectedValue
            Case "KEY_BIDCASE"
                Dim sJS As String = String.Concat("'", dr("KBSID"), "','", dr("KBID"), "','", dr("KBNAME"), "','", dr("KBDESC1"), "'")
                sJS &= String.Concat(",'", dr("MUSTFILL"), "','", dr("ORGKINDGW"), "','", dr("KSORT"), "','", dr("USELATESTVER"), "'")
                sJS &= String.Concat(",'", dr("DOWNLOADRPT"), "','", dr("RPTNAME"), "','", dr("UPLOADFL1"), "','", dr("SENTBATVER"), "'")
                sJS &= String.Concat(",'", dr("USEMEMO1"), "'")
                'sJS &= String.Concat(",'", dr("USEMEMO1"), "','", dr("DataGrid08"), "'")
                NewNode.NavigateUrl = String.Concat("javascript:returnValue17(", sJS, ");")
            Case "SYS_BASICSALARY"
                'returnValue16(v_sbsid, v_swdate, v_monthwage, v_hourwage)
                NewNode.NavigateUrl = String.Concat("javascript:returnValue16('", dr("SBSID"), "','", dr("SWDATE"), "','", dr("MONTHLYWAGE"), "','", dr("HOURLYWAGE"), "');")
            Case "SYS_VAR"
                NewNode.NavigateUrl = String.Concat("javascript:returnValue15('", dr("KID"), "','", dr("Name"), "','", dr("MEMO1"), "');")
            Case "Key_TrainType1"
                NewNode.NavigateUrl = "javascript:returnValue('" & dr("KID") & "','" & dr("Name") & "','0','0','','','','','','','','','','');"
            Case "Key_TrainType2"
                NewNode.NavigateUrl = "javascript:returnValue('" & dr("KID") & "','" & dr("Name") & "','1','" & dr("Parent") & "','','','','','','','','','','');"
            Case "Key_TrainType3"
                NewNode.NavigateUrl = "javascript:returnValue('" & dr("KID") & "','" & dr("Name") & "','2','" & dr("Parent") & "','','','','','','','','','','');"
            Case "Key_Leave"
                NewNode.NavigateUrl = String.Concat("javascript:returnValueKL('", dr("KID"), "','", dr("Name"), "','", dr("MinusPoint"), "','", dr("NOUSE"), "','", dr("LEAVESORT"), "','", dr("ENGNAME"), "');")
            Case "Key_Sanction"
                NewNode.NavigateUrl = "javascript:returnValue('" & dr("KID") & "','" & dr("Name") & "','','','','" & dr("AddMinus") & "','" & dr("Point") & "','','','','','','','');"
            Case "Key_Plan"
                Dim V_QuesID As String = If(V_sPath = "TIMS", Convert.ToString(dr("QID")), If(V_sPath = "DEMO", Convert.ToString(dr("QuesID")), Convert.ToString(dr("QID"))))
                Dim V_Value8 As String = String.Concat("'", dr("KID"), "','", dr("NAME"), "','", dr("PlanType"), "','", dr("IsOnLine"), "','", dr("ClsYear"), "','", V_QuesID, "','", dr("EmailSend"), "','", dr("BlackList"), "','", dr("QueryDisplay"), "','", dr("IdentityID"), "','", dr("Reusable"), "','", dr("useECFA"), "','", dr("PropertyID"), "'")
                'SBK2TPLANID /SBLACK2TPLANID
                V_Value8 &= String.Concat(",'", dr("SBLACKTYPE"), "','", dr("SBK2TPLANID"), "'")
                NewNode.NavigateUrl = String.Concat("javascript:returnValue8(", V_Value8, ");")
            Case "Key_CostItem"
                NewNode.NavigateUrl = "javascript:returnValue11('" & dr("KID") & "','" & dr("Name") & "','" & dr("Sort").ToString & "','" & dr("ItemageName") & "','" & dr("ItemCostName") & "');"
            Case "Key_CostItem2"
                NewNode.NavigateUrl = "javascript:returnValue11('" & dr("KID") & "','" & dr("Name") & "','" & dr("Sort").ToString & "','','" & dr("ItemCostName") & "');"
            Case "Key_DGTHour"
                NewNode.NavigateUrl = "javascript:returnValue('" & dr("KID") & "','" & dr("Name") & "','','','','','','','','','" & dr("DGHour").ToString & "','','','');"
            Case "Key_Degree"
                NewNode.NavigateUrl = "javascript:returnValue('" & dr("KID") & "','" & dr("Name") & "','','','','','','','','" & dr("Sort").ToString & "','','','','" & dr("DegreeType").ToString & "');"
            Case "Key_Identity"
                s_TMP_URL_1 = "'" & dr("KID") & "','" & dr("Name") & "','" & dr("UnUsedYear").ToString & "','" & dr("MergeID").ToString & "'," & If(Convert.ToString(dr("Subsidy")) = "Y", "true", "false") & ",'" & dr("SORT28").ToString & "','" & dr("SUPPLYID").ToString & "','" & dr("NOSHOWMI").ToString & "'"
                NewNode.NavigateUrl = "javascript:returnValue13(" & s_TMP_URL_1 & ");"
            Case "Key_TableMgr"
                NewNode.NavigateUrl = "javascript:returnValue40('" & dr("KID") & "','" & dr("Name") & "','" & dr("KeyTable") & "');"
            Case "Key_JobGroup" '職群代碼
                NewNode.NavigateUrl = "javascript:returnValue11('" & dr("KID") & "','" & dr("Name") & "','" & dr("Sort").ToString & "','','');"
            Case Else
                NewNode.NavigateUrl = "javascript:returnValue('" & dr("KID") & "','" & dr("Name") & "','','','','','','','','','','','','');"
        End Select
        NewNode.Target = "mainFrame"

        If ParentNode Is Nothing Then
            objTreeView.Nodes.Add(NewNode)
        Else
            'ParentNode.Nodes.Add(NewNode)
            ParentNode.ChildNodes.Add(NewNode)
        End If

#Region "(目前未使用)"

        ''加入子節點
        'Dim strRid As String = dr("RID") & "/"
        'strFilter = "Relship like '%" & strRid & "%'"        '先找出符合父節點 xxx\ 開頭的關係
        'For Each drChild In objTable.Select(strFilter)
        '    Dim strRelship As String = drChild("Relship")
        '    Dim pos As Integer = strRelship.IndexOf(strRid)

        '    '若出現格式為「%父節點/子節點/」的Relship值，視為子節點
        '    If pos <> -1 And (pos + strRid.Length) < strRelship.Length Then
        '        If strRelship.IndexOf("/", pos + strRid.Length) = strRelship.Length - 1 Then
        '            AddTreeNodes(drChild, objTable, objTreeView, NewNode)
        '        End If
        '    End If
        'Next

#End Region
    End Sub

    '修改 (id已建立)
    Private Sub CmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
        Update_Key_Table()
    End Sub

    ''' <summary>修改 (id已建立) </summary>
    Sub Update_Key_Table()
        keycode.Text = TIMS.ClearSQM(keycode.Text)
        keyname.Text = TIMS.ClearSQM(keyname.Text)

        Dim v_KeyType As String = TIMS.GetListValue(KeyType)
        Dim s_mytable As String = v_KeyType 'KeyType.SelectedValue
        Dim IDStr As String = ""    '資料庫的主鍵欄位
        Dim NameStr As String = ""  '資料庫的名稱欄位
        Dim TMDISearch As String = ""

        Select Case v_KeyType 'KeyType.SelectedValue
            Case "SYS_BASICSALARY" '(排除)
                hid_BASICSALARY_SBSID.Value = TIMS.ClearSQM(hid_BASICSALARY_SBSID.Value)
                keycode.Text = hid_BASICSALARY_SBSID.Value
                SWDATE.Text = TIMS.Cdate3(SWDATE.Text)
                If SWDATE.Text = "" OrElse hid_BASICSALARY_SBSID.Value = "" Then
                    Common.MessageBox(Page, "未輸入實施日期(鍵值代碼)，請確認!!")
                    Return
                End If
                Dim dtBs As New DataTable
                Dim s_sql As String = "SELECT 1 FROM SYS_BASICSALARY WHERE SWDATE=@SWDATE AND SBSID!=@SBSID"
                Dim sCmd As New SqlCommand(s_sql, objconn)
                With sCmd
                    .Parameters.Clear()
                    .Parameters.Add("SWDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(SWDATE.Text)
                    .Parameters.Add("SBSID", SqlDbType.BigInt).Value = Val(hid_BASICSALARY_SBSID.Value)
                    dtBs.Load(.ExecuteReader())
                End With
                If dtBs.Rows.Count > 0 Then
                    Common.MessageBox(Page, "此 實施日期(鍵值代碼)已存在" + keycode.Text)
                    Return
                End If

            Case Else
                If keycode.Text = "" Then
                    Common.MessageBox(Page, "未輸入鍵值代碼，請確認!!")
                    Exit Sub
                End If

        End Select

        Select Case s_mytable
            Case "KEY_BIDCASE"
                IDStr = "KBSID"
                NameStr = "KBNAME"
            Case "SYS_BASICSALARY"
                keycode.Text = hid_BASICSALARY_SBSID.Value
                IDStr = "SBSID"
                NameStr = "SWDATE"
            Case "SYS_VAR"
                IDStr = "ITEMNAME"
                NameStr = "ITEMVALUE"
            Case "ID_DISTRICT"
                IDStr = "DISTID"
                NameStr = "Name"
            Case "Key_OrgType"
                IDStr = "OrgTypeID"
                NameStr = "Name"
            Case "Key_Degree"
                IDStr = "DegreeID"
                NameStr = "Name"
            Case "Key_Identity"
                IDStr = "IdentityID"
                NameStr = "Name"
            Case "Key_Military"
                IDStr = "MilitaryID"
                NameStr = "Name"
            Case "Key_RejectTReason"
                IDStr = "RTReasonID"
                NameStr = "Reason"
            Case "Key_SelResult"
                IDStr = "SelResultID"
                NameStr = "Name"
            Case "Key_Subsidy"
                IDStr = "SubsidyID"
                NameStr = "Name"
            Case "Key_TtrainingSource"
                IDStr = "SourceID"
                NameStr = "Name"
            Case "Key_GradState"
                IDStr = "GradID"
                NameStr = "Name"
            Case "Key_JoblessWeek"
                IDStr = "JoblessID"
                NameStr = "Name"
            Case "Key_HourRan"
                IDStr = "HRID"
                NameStr = "HourRanName"
            Case "Key_TrainExp"
                IDStr = "TEID"
                NameStr = "TrainExpName"
            Case "Key_HandicatType"
                IDStr = "HandTypeID"
                NameStr = "Name"
            Case "Key_HandicatLevel"
                IDStr = "HandLevelID"
                NameStr = "Name"
            Case "Key_HandicatType2"
                IDStr = "HandTypeID2"
                NameStr = "Name"
            Case "Key_HandicatLevel2"
                IDStr = "HandLevelID2"
                NameStr = "Name"
            Case "Key_Plan"   '計畫
                IDStr = "TPlanID"
                NameStr = "PlanName"
            Case "Key_Leave"
                IDStr = "LeaveID"
                NameStr = "Name"
            Case "Key_Sanction"
                IDStr = "SanID"
                NameStr = "Name"
            Case "Key_Exam"
                IDStr = "ExamID"
                NameStr = "Name"
            Case "Key_TrainType1"
                s_mytable = "Key_TrainType"
                IDStr = "BusID"
                NameStr = "BusName"
            Case "Key_TrainType2"
                s_mytable = "Key_TrainType"
                IDStr = "JobID"
                NameStr = "JobName"
                TMDISearch = String.Concat(" and Parent='", Parent1.Value, "'")
            Case "Key_TrainType3"
                s_mytable = "Key_TrainType"
                IDStr = "TrainID"
                NameStr = "TrainName"
                TMDISearch = String.Concat(" and Parent='", Parent1.Value, "'")
            Case "Key_CostItem"
                IDStr = "CostID"
                NameStr = "CostName"
            Case "Key_CostItem2"
                IDStr = "CostID"
                NameStr = "CostName"
            Case "Key_Budget"
                IDStr = "BudID"
                NameStr = "BudName"
            Case "Key_DGTHour"
                IDStr = "DGID"
                NameStr = "DGName"
            Case "Key_Salary"
                IDStr = "SalID"
                NameStr = "SalName"
            Case "Key_GetJob"
                IDStr = "GJID"
                NameStr = "GJName"
            Case "Key_NonGetJob"
                IDStr = "NGJID"
                NameStr = "NGJName"
            Case "Key_BusVisitCase"
                IDStr = "BVCID"
                NameStr = "BVCName"
            Case "Key_AgeRange"
                IDStr = "ARID"
                NameStr = "ARName"
            Case "Key_WorkYear"
                IDStr = "WYID"
                NameStr = "WYName"
            Case "Key_Emp"
                IDStr = "KEID"
                NameStr = "KEName"
'Case "Key_KindofJob"
'    IDStr = "KEID"
'    NameStr = "KEName"
'Case "Key_ProSkill"
'    IDStr = "KEID"
'    NameStr = "KEName"
            Case "Key_Trade"
                IDStr = "TradeID"
                NameStr = "TradeName"
            Case "Key_NotOpenReason"
                IDStr = "NORID"
                NameStr = "NORName"
            Case "Key_ClassCatelog"
                IDStr = "CCID"
                NameStr = "CCName"
            Case "Key_CheckRan"
                IDStr = "CHID"
                NameStr = "CheckRanName"
            Case "ID_Invest"
                IDStr = "IVID"
                NameStr = "InvestName"
            Case "OB_Funds"
                IDStr = "FID"
                NameStr = "FName"
            Case "Key_Native"
                IDStr = "KNID"
                NameStr = "Name"
            Case "Key_SurveyType"
                IDStr = "STID"
                NameStr = "STName"
            Case "Key_JobGroup" '職群鍵詞檔
                IDStr = "JGID"
                NameStr = "JGNAME"
            Case "Key_EnterPoint" '錄訓百分比代碼
                IDStr = "KID"
                NameStr = "KNAME"
            Case "Key_TableMgr"
                IDStr = "SerNum"
                NameStr = "KeyType"
        End Select

        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Select Case s_mytable
            Case "KEY_BIDCASE"
                sql = String.Concat("SELECT * FROM ", s_mytable, " WHERE ", IDStr, "='", hid_BIDCASE_KBSID.Value, "'")
            Case Else
                sql = String.Concat("SELECT * FROM ", s_mytable, " WHERE ", IDStr, "='", keycode.Text, "' ", TMDISearch)
        End Select
        Try
            dt = DbAccess.GetDataTable(sql, da, objconn)
        Catch ex As Exception
            Common.MessageBox(Page, String.Concat("查詢有誤，", ex.Message))
            Return
        End Try

        If dt.Rows.Count = 0 Then
            Common.MessageBox(Page, String.Concat("此鍵值代碼(", keycode.Text, ")不存在,請新增"))
            Page.RegisterStartupScript("window_onload", "<script language='javascript'>display_Item();</script>")
            'btnSearch_Click(sender, e)
            SbSearch1()
            Exit Sub
        End If

        Dim V_sPath As String = Convert.ToString(ViewState("sPath")) 'Convert.ToString(dr("QID"))
        Dim dr As DataRow = dt.Rows(0)
        Select Case s_mytable '基礎值
            Case "SYS_BASICSALARY" '(排除)-基礎值
            Case Else
                dr(NameStr) = keyname.Text
        End Select

        Select Case s_mytable '附加值
            Case "KEY_BIDCASE"
                Dim v_RBL_ORGKINDGW As String = TIMS.GetListValue(RBL_ORGKINDGW)
                txt_KSORT.Text = TIMS.ClearSQM(txt_KSORT.Text)
                txt_RPTNAME.Text = TIMS.ClearSQM(txt_RPTNAME.Text)

                'KBSID／dr(NameStr) = keyname.Text KBNAME
                dr("KBID") = TIMS.ClearSQM(keycode.Text)
                dr("MUSTFILL") = If(CB_MUSTFILL.Checked, "Y", Convert.DBNull)
                dr("USELATESTVER") = If(cb_USELATESTVER.Checked, "Y", Convert.DBNull)
                dr("DOWNLOADRPT") = If(cb_DOWNLOADRPT.Checked, "Y", Convert.DBNull)
                dr("UPLOADFL1") = If(cb_UPLOADFL1.Checked, "Y", Convert.DBNull)
                dr("SENTBATVER") = If(cb_SENTBATVER.Checked, "Y", Convert.DBNull)
                dr("USEMEMO1") = If(cb_USEMEMO1.Checked, "Y", Convert.DBNull)
                'dr("DataGrid08") = If(cb_DataGrid08.Checked, "Y", Convert.DBNull)
                dr("ORGKINDGW") = If(v_RBL_ORGKINDGW = "W", v_RBL_ORGKINDGW, If(v_RBL_ORGKINDGW = "G", v_RBL_ORGKINDGW, Convert.DBNull))
                dr("KSORT") = If(txt_KSORT.Text <> "", Val(txt_KSORT.Text), Convert.DBNull)
                dr("RPTNAME") = If(txt_RPTNAME.Text <> "", txt_RPTNAME.Text, Convert.DBNull) 'TIMS.ClearSQM(txt_RPTNAME.Text)
                dr("KBDESC1") = If(txt_KBDESC1.Text <> "", txt_KBDESC1.Text, Convert.DBNull)

            Case "SYS_BASICSALARY" '基本工資 https://www.mol.gov.tw/1607/28162/28166/28180/28182/
                SWDATE.Text = TIMS.Cdate3(SWDATE.Text)
                dr("MONTHLYWAGE") = TIMS.ClearSQM(MONTHLYWAGE.Text)
                dr("HOURLYWAGE") = TIMS.ClearSQM(HOURLYWAGE.Text)
                dr("SWDATE") = TIMS.Cdate2(SWDATE.Text)
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now

            Case "SYS_VAR"
                dr("MEMO1") = TIMS.ClearSQM(txtmemo1.Text)
            Case "Key_Degree"
                dr("DegreeType") = rblDegreeType.SelectedValue
                dr("Sort") = Sort.Text
            Case "Key_JobGroup" '職群鍵詞檔
                Sort.Text = TIMS.ClearSQM(Sort.Text)
                dr("Sort") = If((Sort.Text <> ""), CInt(Val(Sort.Text)), Convert.DBNull) ' Sort.Text
            Case "Key_Plan"
                dr("IsOnLine") = IsOnLine.SelectedValue
                dr("PlanType") = If(TypeValue.Value <> "", Val(TypeValue.Value), Convert.DBNull)
                dr("ClsYear") = If(ClsYear.Text <> "", ClsYear.Text, Convert.DBNull)
                'TIMS: 欄位功能取消，放到 Plan_Questionary
                dr("QuesID") = If(V_sPath = "TIMS", Convert.DBNull, TIMS.GetListValue(QuesType))
                dr("EmailSend") = If(rblEmailSend.SelectedValue <> "", rblEmailSend.SelectedValue, "Y")
                dr("Reusable") = If(rblReusable.SelectedValue <> "", rblReusable.SelectedValue, "N")
                dr("BlackList") = If(rblBlackList.SelectedValue <> "", rblBlackList.SelectedValue, "N")

                Dim v_SBLACKTYPE As String = TIMS.GetListValue(rblSBLACKTYPE)
                dr("SBLACKTYPE") = If(v_SBLACKTYPE <> "", v_SBLACKTYPE, Convert.DBNull)
                Dim v_SBLACK2TPLANID As String = TIMS.GetCblValue(cblSBK2TPLANID)
                dr("SBLACK2TPLANID") = If(v_SBLACK2TPLANID <> "", v_SBLACK2TPLANID, Convert.DBNull)

                dr("QueryDisplay") = If(rblQueryDisplay.SelectedValue <> "", rblQueryDisplay.SelectedValue, "N")
                '是否適用於ECFA協助基金
                dr("useECFA") = If(rbluseECFA.SelectedValue <> "", rbluseECFA.SelectedValue, "N")
                dr("PropertyID") = If(ddlPropertyID.SelectedValue <> "N", Val(ddlPropertyID.SelectedValue), Convert.DBNull)
            Case "Key_Leave"
                MinusPoint.Text = TIMS.ClearSQM(MinusPoint.Text)
                Sort.Text = TIMS.ClearSQM(Sort.Text)
                EngkeyName.Text = TIMS.ClearSQM(EngkeyName.Text)
                dr("MinusPoint") = MinusPoint.Text
                dr("NOUSE") = If(cb_LEAVE_NOUSE.Checked, "Y", Convert.DBNull)
                dr("LEAVESORT") = If(Sort.Text <> "", Sort.Text, Convert.DBNull) 'TIMS.GetValue1(Sort.Text)
                dr("ENGNAME") = If(EngkeyName.Text <> "", EngkeyName.Text, Convert.DBNull)
            Case "Key_Sanction"
                dr("AddMinus") = AddMinus.Text
                dr("Point") = point.Text
            Case "Key_CostItem"
                dr("Sort") = Sort.Text
                dr("ItemageName") = txtItemageName.Text
                dr("ItemCostName") = txtItemCostName.Text
            Case "Key_CostItem2"
                dr("Sort") = Sort.Text
                dr("ItemCostName") = txtItemCostName.Text
            Case "Key_Identity"
                '可申請生活津貼
                dr("Subsidy") = If(chkSubsidy.Checked, "Y", Convert.DBNull)
                '停用年度
                Dim v_ddlUnUsedYear As String = TIMS.GetListValue(ddlUnUsedYear)
                dr("UnUsedYear") = If(v_ddlUnUsedYear <> "", v_ddlUnUsedYear, Convert.DBNull)
                '併入身分別
                Dim v_ddlMergeID As String = TIMS.GetListValue(ddlMergeID)
                dr("MergeID") = If(v_ddlMergeID <> "", v_ddlMergeID, Convert.DBNull)
                '排序序號
                txtSORT28.Text = TIMS.ClearSQM(txtSORT28.Text)
                dr("SORT28") = If(txtSORT28.Text <> "", txtSORT28.Text, Convert.DBNull)
                '補助比例
                Dim v_rblSUPPLYID As String = TIMS.GetListValue(rblSUPPLYID)
                dr("SUPPLYID") = If(v_rblSUPPLYID <> "", v_rblSUPPLYID, Convert.DBNull)
                '主要參訓身分別不顯示
                Dim v_rblNOSHOWMI As String = TIMS.GetListValue(rblNOSHOWMI)
                dr("NOSHOWMI") = If(v_rblNOSHOWMI <> "", v_rblNOSHOWMI, Convert.DBNull)
            Case "Key_TableMgr"
                dr("KeyTable") = txtKeyTable.Text
            Case "Key_DGTHour"
                dr("DGHour") = DGHour.Text
            Case "ID_DISTRICT", "ID_Invest", "OB_Funds"
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now
        End Select
        DbAccess.UpdateDataTable(dt, da)

        '附加值/'附加動作
        Select Case s_mytable
            Case "SYS_VAR"
                'Call Utl_ClearDirConfig()
                'Dim t_VAL As String = TIMS.Utl_GetConfigVAL(objconn, keycode.Text, 1)
                keyname.Text = TIMS.Utl_GetConfigVAL(objconn, keycode.Text, 1) 't_VAL 
            Case "Key_Plan"
                Dim v_QuesType As String = TIMS.GetListValue(QuesType)
                If V_sPath = "TIMS" Then
                    Dim HHt As New Hashtable
                    TIMS.SetMyValue2(HHt, "keycode", keycode.Text)
                    TIMS.SetMyValue2(HHt, "QuesType", v_QuesType)
                    Call UPD_PLAN_QUESTIONARY(HHt, objconn)
                End If
                'Update Plan_Identity
                Call UPDATE_PLAN_IDENTITY(sm, IdentityID, keycode.Text, objconn)
        End Select

        'btnSearch_Click(sender, e)
        SbSearch1()

        Common.MessageBox(Me, "修改成功!")

    End Sub

    ''' <summary> 新增 (id未建立) </summary>
    Sub Insert_Key_Table()
        keycode.Text = TIMS.ClearSQM(keycode.Text)
        keyname.Text = TIMS.ClearSQM(keyname.Text)

        Dim v_KeyType As String = TIMS.GetListValue(KeyType)
        Select Case v_KeyType 'KeyType.SelectedValue
            Case "SYS_BASICSALARY" '(排除)
                SWDATE.Text = TIMS.Cdate3(SWDATE.Text)
                If SWDATE.Text = "" Then
                    Common.MessageBox(Page, "未輸入實施日期(鍵值代碼)，請確認!!")
                    Return
                End If
                Dim dtBs As New DataTable
                Dim s_sql As String = "SELECT 1 FROM SYS_BASICSALARY WHERE SWDATE=@SWDATE"
                Dim sCmd As New SqlCommand(s_sql, objconn)
                With sCmd
                    .Parameters.Clear()
                    .Parameters.Add("SWDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(SWDATE.Text)
                    dtBs.Load(.ExecuteReader())
                End With
                If dtBs.Rows.Count > 0 Then
                    Common.MessageBox(Page, "此 實施日期(鍵值代碼)已存在" + keycode.Text)
                    Return
                End If

            Case Else
                If keycode.Text = "" Then
                    Common.MessageBox(Page, "未輸入鍵值代碼，請確認!!")
                    Exit Sub
                End If
        End Select

        Dim sqlAdapter As SqlDataAdapter = Nothing
        Dim sqlTable As DataTable = Nothing
        Dim sqldr As DataRow = Nothing
        Dim StrSql As String = ""

        Dim V_sPath As String = Convert.ToString(ViewState("sPath")) 'Convert.ToString(dr("QID"))
        Select Case v_KeyType 'KeyType.SelectedValue
            Case "SYS_BASICSALARY"
                'SWDATE.Text = TIMS.cdate3(SWDATE.Text)
                Dim iSBSID As Integer = DbAccess.GetNewId(objconn, "SYS_BASICSALARY_SBSID_SEQ,SYS_BASICSALARY,SBSID")
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("SBSID", iSBSID)
                parms.Add("MONTHLYWAGE", Val(MONTHLYWAGE.Text))
                parms.Add("HOURLYWAGE", Val(HOURLYWAGE.Text))
                parms.Add("SWDATE", TIMS.Cdate2(SWDATE.Text))
                parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                Dim tSql As String = "INSERT INTO SYS_BASICSALARY(SBSID,MONTHLYWAGE,HOURLYWAGE,SWDATE,MODIFYACCT,MODIFYDATE) VALUES (@SBSID,@MONTHLYWAGE,@HOURLYWAGE,@SWDATE,@MODIFYACCT,GETDATE())"
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)
            Case "SYS_VAR"
                StrSql = "SELECT SVID,ITEMNAME,ITEMVALUE FROM SYS_VAR WHERE SPAGE='Config' AND UPPER(ITEMNAME)=UPPER('" & keycode.Text & "')"
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                '不存在,新增記錄
                Dim iSVID As Integer = DbAccess.GetNewId(objconn, "SYS_VAR_SVID_SEQ,SYS_VAR,SVID")
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("SVID") = iSVID
                sqldr("ITEMNAME") = keycode.Text
                sqldr("ITEMVALUE") = keyname.Text
                sqldr("MEMO1") = txtmemo1.Text
                sqldr("SPAGE") = "Config"
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)
                'Call Utl_ClearDirConfig()
            Case "ID_DISTRICT"
                StrSql = "SELECT DISTID ,Name From ID_DISTRICT WHERE DISTID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                '不存在,新增記錄
                sqldr = DbAccess.GetInsertRow("ID_DISTRICT", sqlTable, sqlAdapter)
                sqldr("DISTID") = keycode.Text
                sqldr("Name") = keyname.Text
                sqldr("ModifyAcct") = sm.UserInfo.UserID
                sqldr("ModifyDate") = Now
                sqlTable.Rows.Add(sqldr)
                'sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)"

                Dim tSql As String = "INSERT INTO ID_DISTRICT (DISTID, NAME, MODIFYACCT, MODIFYDATE) "
                tSql &= " VALUES(@DISTID, @NAME, @MODIFYACCT, GETDATE()) "
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("DISTID", keycode.Text)
                parms.Add("NAME", keyname.Text)
                parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)

            Case "Key_OrgType"
                '判斷ID是否存在
                StrSql = "SELECT OrgTypeID,Name From Key_OrgType Where OrgTypeID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                '不存在,新增記錄
                sqldr = DbAccess.GetInsertRow("Key_OrgType", sqlTable, sqlAdapter)
                sqldr("OrgTypeID") = keycode.Text
                sqldr("Name") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                'sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)"

                Dim tSql As String = "INSERT INTO Key_OrgType (ORGTYPEID, NAME) "
                tSql &= " VALUES(@ORGTYPEID, @NAME) "
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("ORGTYPEID", keycode.Text)
                parms.Add("NAME", keyname.Text)
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)

            Case "Key_Degree"
                StrSql = "SELECT DegreeID,Name,DegreeType,Sort FROM Key_Degree Where DegreeID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow("Key_Degree", sqlTable, sqlAdapter)
                sqldr("DegreeID") = keycode.Text
                sqldr("Name") = keyname.Text
                sqldr("DegreeType") = rblDegreeType.SelectedValue
                sqldr("Sort") = Sort.Text
                sqlTable.Rows.Add(sqldr)
                'sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)"

                Dim tSql As String = "INSERT INTO Key_Degree (DEGREEID, NAME, DEGREETYPE, SORT) "
                tSql &= " VALUES(@DEGREEID, @NAME, @DEGREETYPE, @SORT) "
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("DEGREEID", keycode.Text)
                parms.Add("NAME", keyname.Text)
                parms.Add("DEGREETYPE", rblDegreeType.SelectedValue)
                parms.Add("SORT", Sort.Text)
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)

            Case "Key_Identity"
                StrSql = "SELECT IdentityID,Name FROM Key_Identity WHERE IdentityID='" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If

                '停用年度
                Dim v_ddlUnUsedYear As String = TIMS.GetListValue(ddlUnUsedYear)
                '併入身分別
                Dim v_ddlMergeID As String = TIMS.GetListValue(ddlMergeID)
                '排序序號
                txtSORT28.Text = TIMS.ClearSQM(txtSORT28.Text)
                '補助比例
                Dim v_rblSUPPLYID As String = TIMS.GetListValue(rblSUPPLYID)
                '主要參訓身分別不顯示
                Dim v_rblNOSHOWMI As String = TIMS.GetListValue(rblNOSHOWMI)

                sqldr = DbAccess.GetInsertRow("Key_Identity", sqlTable, sqlAdapter)
                sqldr("IdentityID") = keycode.Text
                sqldr("Name") = keyname.Text
                '可申請生活津貼
                sqldr("Subsidy") = If(chkSubsidy.Checked, "Y", Convert.DBNull)
                sqldr("UnUsedYear") = If(v_ddlUnUsedYear <> "", v_ddlUnUsedYear, Convert.DBNull)
                sqldr("MergeID") = If(v_ddlMergeID <> "", v_ddlMergeID, Convert.DBNull)
                sqldr("SORT28") = If(txtSORT28.Text <> "", txtSORT28.Text, Convert.DBNull)
                sqldr("SUPPLYID") = If(v_rblSUPPLYID <> "", v_rblSUPPLYID, Convert.DBNull)
                sqldr("NOSHOWMI") = If(v_rblNOSHOWMI <> "", v_rblNOSHOWMI, Convert.DBNull)
                'sqlTable.Rows.Add(sqldr)
                'sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)"

                Dim tSql As String = ""
                tSql = ""
                tSql &= " INSERT INTO KEY_IDENTITY (IDENTITYID, NAME, UNUSEDYEAR, MERGEID, SUBSIDY"
                tSql &= " ,SORT28,SUPPLYID,NOSHOWMI) "
                tSql &= " VALUES(@IDENTITYID, @NAME, @UNUSEDYEAR, @MERGEID, @SUBSIDY"
                tSql &= " ,@SORT28,@SUPPLYID,@NOSHOWMI) "

                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("IDENTITYID", keycode.Text)
                parms.Add("NAME", keyname.Text)
                parms.Add("UNUSEDYEAR", If(v_ddlUnUsedYear <> "", v_ddlUnUsedYear, Convert.DBNull))
                parms.Add("MERGEID", If(v_ddlMergeID <> "", v_ddlMergeID, Convert.DBNull))
                '可申請生活津貼
                parms.Add("SUBSIDY", If(chkSubsidy.Checked, "Y", Convert.DBNull))
                parms.Add("SORT28", If(txtSORT28.Text <> "", txtSORT28.Text, Convert.DBNull))
                parms.Add("SUPPLYID", If(v_rblSUPPLYID <> "", v_rblSUPPLYID, Convert.DBNull))
                parms.Add("NOSHOWMI", If(v_rblNOSHOWMI <> "", v_rblNOSHOWMI, Convert.DBNull))
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)

            Case "Key_Military"
                StrSql = "SELECT MilitaryID,Name FROM Key_Military WHERE MilitaryID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow("Key_Military", sqlTable, sqlAdapter)
                sqldr("MilitaryID") = keycode.Text
                sqldr("Name") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                'sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)"

                Dim tSql As String = "INSERT INTO Key_Military (MILITARYID, NAME) "
                tSql &= " VALUES(@MILITARYID, @NAME) "
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("MILITARYID", keycode.Text)
                parms.Add("NAME", keyname.Text)
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)

            Case "Key_Plan"  '20080724 Andy「QuesID」 欄位目前不使用；改用ID_Questionary 、Plan_Questionary資料表，請參修改功能 
                '新增。
                StrSql = "SELECT TPlanID,PlanName,IsOnLine,PlanType,ClsYear FROM KEY_PLAN where TPlanID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow("Key_Plan", sqlTable, sqlAdapter)
                sqldr("TPlanID") = keycode.Text
                sqldr("PlanName") = keyname.Text
                sqldr("IsOnLine") = IsOnLine.SelectedValue
                sqldr("PlanType") = If(TypeValue.Value <> "", TypeValue.Value, Convert.DBNull)
                sqldr("ClsYear") = If(ClsYear.Text <> "", ClsYear.Text, Convert.DBNull)
                'TIMS: 欄位功能取消，放到 Plan_Questionary
                sqldr("QuesID") = If(V_sPath = "TIMS", Convert.DBNull, TIMS.GetListValue(QuesType))

                sqldr("EmailSend") = If(rblEmailSend.SelectedValue <> "", rblEmailSend.SelectedValue, "Y")
                sqldr("Reusable") = If(rblReusable.SelectedValue <> "", rblReusable.SelectedValue, "N")
                sqldr("BlackList") = If(rblBlackList.SelectedValue <> "", rblBlackList.SelectedValue, "N")

                Dim v_SBLACKTYPE As String = TIMS.GetListValue(rblSBLACKTYPE)
                sqldr("SBLACKTYPE") = If(v_SBLACKTYPE <> "", v_SBLACKTYPE, Convert.DBNull)
                Dim v_cblSBK2TPLANID As String = TIMS.GetCblValue(cblSBK2TPLANID)
                sqldr("SBLACK2TPLANID") = If(v_cblSBK2TPLANID <> "", v_cblSBK2TPLANID, Convert.DBNull)

                sqldr("QueryDisplay") = If(rblQueryDisplay.SelectedValue <> "", rblQueryDisplay.SelectedValue, "N")
                '是否適用於ECFA協助基金
                sqldr("useECFA") = If(rbluseECFA.SelectedValue <> "", rbluseECFA.SelectedValue, "N")
                sqldr("PropertyID") = If(ddlPropertyID.SelectedValue <> "N", Val(ddlPropertyID.SelectedValue), Convert.DBNull)
                sqlTable.Rows.Add(sqldr)
                'sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)"

                Dim tSql As String = ""
                tSql = ""
                tSql &= " INSERT INTO KEY_PLAN (TPLANID, PLANNAME, ISONLINE, PLANTYPE, PLANSNAME, CLSYEAR, PUBPRINT"
                tSql &= " ,QUESID, EMAILSEND, REUSABLE, BLACKLIST,SBLACKTYPE,SBLACK2TPLANID"
                tSql &= " ,QUERYDISPLAY, USEECFA, PROPERTYID) "
                tSql &= " VALUES(@TPLANID, @PLANNAME, @ISONLINE, @PLANTYPE, @PLANSNAME, @CLSYEAR, @PUBPRINT"
                tSql &= " ,@QUESID, @EMAILSEND, @REUSABLE, @BLACKLIST,@SBLACKTYPE,@SBLACK2TPLANID"
                tSql &= " ,@QUERYDISPLAY, @USEECFA, @PROPERTYID) "

                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("TPLANID", sqldr("TPlanID"))
                parms.Add("PLANNAME", sqldr("PlanName"))
                parms.Add("ISONLINE", sqldr("IsOnLine"))
                parms.Add("PLANTYPE", Convert.DBNull)
                parms.Add("PLANSNAME", Convert.DBNull)
                parms.Add("CLSYEAR", sqldr("ClsYear"))
                parms.Add("PUBPRINT", Convert.DBNull)
                parms.Add("QUESID", sqldr("QuesID"))
                parms.Add("EMAILSEND", sqldr("EmailSend"))
                parms.Add("REUSABLE", sqldr("Reusable"))
                parms.Add("BLACKLIST", sqldr("BlackList"))
                parms.Add("SBLACKTYPE", sqldr("SBLACKTYPE"))
                parms.Add("SBLACK2TPLANID", sqldr("SBLACK2TPLANID"))
                '--update Key_Plan set SBLACKTYPE =1 where SBLACKTYPE Is null
                '--update Key_Plan set SBLACK2TPLANID =TPLANID where SBLACK2TPLANID Is null
                parms.Add("QUERYDISPLAY", sqldr("QueryDisplay"))
                parms.Add("USEECFA", sqldr("useECFA"))
                parms.Add("PROPERTYID", sqldr("PropertyID"))
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)

                Dim v_QuesType As String = TIMS.GetListValue(QuesType)
                If V_sPath = "TIMS" Then
                    Dim HHt As New Hashtable
                    TIMS.SetMyValue2(HHt, "keycode", keycode.Text)
                    TIMS.SetMyValue2(HHt, "QuesType", v_QuesType)
                    Call UPD_PLAN_QUESTIONARY(HHt, objconn)
                End If
                'Update Plan_Identity
                Call UPDATE_PLAN_IDENTITY(sm, IdentityID, keycode.Text, objconn)

            Case "Key_RejectTReason"
                StrSql = "SELECT RTReasonID,Reason FROM Key_RejectTReason WHERE RTReasonID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow("Key_RejectTReason", sqlTable, sqlAdapter)
                sqldr("RTReasonID") = keycode.Text
                sqldr("Reason") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                'sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)"

                Dim tSql As String = "INSERT INTO Key_RejectTReason (RTREASONID, REASON) "
                tSql &= " VALUES(@RTREASONID, @REASON) "

                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("RTREASONID", keycode.Text)
                parms.Add("REASON", keyname.Text)
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)

            Case "Key_SelResult"
                StrSql = "SELECT SelResultID,Name From Key_SelResult Where SelResultID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If

                Dim tSql As String = "INSERT INTO Key_SelResult (SELRESULTID, NAME) "
                tSql &= " VALUES(@SELRESULTID, @NAME) "
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("SELRESULTID", keycode.Text)
                parms.Add("NAME", keyname.Text)
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)

            Case "Key_Subsidy"
                StrSql = "SELECT SubsidyID,Name FROM Key_Subsidy WHERE SubsidyID='" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow("Key_Subsidy", sqlTable, sqlAdapter)
                sqldr("SubsidyID") = keycode.Text
                sqldr("Name") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                'sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)"

                Dim tSql As String = "INSERT INTO Key_Subsidy (SUBSIDYID, NAME) "
                tSql &= " VALUES(@SUBSIDYID, @NAME) "

                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("SUBSIDYID", keycode.Text)
                parms.Add("NAME", keyname.Text)
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)

            Case "Key_TtrainingSource"
                StrSql = "SELECT SourceID,Name From Key_TtrainingSource WHERE SourceID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow("Key_TtrainingSource", sqlTable, sqlAdapter)
                sqldr("SourceID") = keycode.Text
                sqldr("Name") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                'sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)"

                Dim tSql As String = "INSERT INTO Key_TtrainingSource (SOURCEID, NAME) "
                tSql &= " VALUES(@SOURCEID, @NAME) "
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("SOURCEID", keycode.Text)
                parms.Add("NAME", keyname.Text)
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)

            Case "Key_TrainType1"
                StrSql = "SELECT BusID,BusName,Levels,Parent FROM Key_TrainType WHERE BusID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow("Key_TrainType", sqlTable, sqlAdapter)
                sqldr("TMID") = DbAccess.GetNewId(objconn, "KEY_TRAINTYPE_TMID_SEQ,KEY_TRAINTYPE,TMID")
                sqldr("BusID") = keycode.Text
                sqldr("BusName") = keyname.Text
                sqldr("Levels") = 0
                sqldr("Parent") = 0
                sqlTable.Rows.Add(sqldr)
                'sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)"

                Dim tSql As String = "INSERT INTO Key_TrainType (TMID, BUSID, BUSNAME, LEVELS, PARENT) "
                tSql &= " VALUES(@TMID, @BUSID, @BUSNAME, @LEVELS, @PARENT) "
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("TMID", sqldr("TMID"))
                parms.Add("BUSID", keycode.Text)
                parms.Add("BUSNAME", keyname.Text)
                parms.Add("LEVELS", sqldr("Levels"))
                parms.Add("PARENT", sqldr("Parent"))
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)

            Case "Key_TrainType2"
                StrSql = "SELECT JobID,JobName,Levels,Parent FROM Key_TrainType WHERE JobID='" & keycode.Text & "' AND Parent='" & Parent1.Value & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow("Key_TrainType", sqlTable, sqlAdapter)
                sqldr("TMID") = DbAccess.GetNewId(objconn, "KEY_TRAINTYPE_TMID_SEQ,KEY_TRAINTYPE,TMID")
                sqldr("JobID") = keycode.Text
                sqldr("JobName") = keyname.Text
                sqldr("Levels") = 1
                sqldr("Parent") = Parent1.Value
                sqlTable.Rows.Add(sqldr)
                'sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)"

                Dim tSql As String = "INSERT INTO Key_TrainType (TMID, JOBID, JOBNAME, LEVELS, PARENT) "
                tSql &= " VALUES(@TMID, @JOBID, @JOBNAME, @LEVELS, @PARENT) "
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("TMID", sqldr("TMID"))
                parms.Add("JOBID", keycode.Text)
                parms.Add("JOBNAME", keyname.Text)
                parms.Add("LEVELS", sqldr("Levels"))
                parms.Add("PARENT", sqldr("Parent"))
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)

            Case "Key_TrainType3"
                StrSql = "SELECT TrainID,TrainName,Levels,Parent FROM Key_TrainType Where TrainID = '" & keycode.Text & "' AND Parent = '" & Parent1.Value & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow("Key_TrainType", sqlTable, sqlAdapter)
                sqldr("TMID") = DbAccess.GetNewId(objconn, "KEY_TRAINTYPE_TMID_SEQ,KEY_TRAINTYPE,TMID")
                sqldr("TrainID") = keycode.Text
                sqldr("TrainName") = keyname.Text
                sqldr("Levels") = 2
                sqldr("Parent") = Parent1.Value
                sqlTable.Rows.Add(sqldr)
                'sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)"

                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("TMID", sqldr("TMID"))
                parms.Add("TRAINID", keycode.Text)
                parms.Add("TRAINNAME", keyname.Text)
                parms.Add("LEVELS", sqldr("Levels"))
                parms.Add("PARENT", sqldr("Parent"))
                Dim tSql As String = "INSERT INTO Key_TrainType (TMID, TRAINID, TRAINNAME, LEVELS, PARENT) "
                tSql &= " VALUES(@TMID, @JOBID, @JOBNAME, @LEVELS, @PARENT) "
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)
            Case "Key_Leave"
                StrSql = "SELECT * From Key_Leave WHERE LeaveID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                '(新增一筆資料到Table裡)"
                MinusPoint.Text = TIMS.ClearSQM(MinusPoint.Text)
                Sort.Text = TIMS.ClearSQM(Sort.Text)
                EngkeyName.Text = TIMS.ClearSQM(EngkeyName.Text)
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("LEAVEID", keycode.Text)
                parms.Add("NAME", keyname.Text)
                parms.Add("MINUSPOINT", If(MinusPoint.Text <> "", MinusPoint.Text, Convert.DBNull))
                parms.Add("NOUSE", If(cb_LEAVE_NOUSE.Checked, "Y", Convert.DBNull))
                parms.Add("LEAVESORT", If(Sort.Text <> "", Sort.Text, Convert.DBNull))
                parms.Add("ENGNAME", If(EngkeyName.Text <> "", EngkeyName.Text, Convert.DBNull))
                Dim tSql As String = ""
                tSql = " INSERT INTO KEY_LEAVE (LEAVEID,NAME,MINUSPOINT,NOUSE,LEAVESORT,ENGNAME) "
                tSql &= " VALUES(@LEAVEID,@NAME,@MINUSPOINT,@NOUSE,@LEAVESORT,@ENGNAME) "
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)
            Case "Key_Sanction"
                StrSql = "SELECT * FROM Key_Sanction WHERE SanID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow("Key_Sanction", sqlTable, sqlAdapter)
                sqldr("SanID") = keycode.Text
                sqldr("Name") = keyname.Text
                sqldr("AddMinus") = AddMinus.Text
                sqldr("Point") = point.Text
                sqlTable.Rows.Add(sqldr)
                'sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)"

                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("SANID", sqldr("SanID"))
                parms.Add("NAME", sqldr("Name"))
                parms.Add("ADDMINUS", sqldr("AddMinus"))
                parms.Add("POINT", sqldr("Point"))
                Dim tSql As String = "INSERT INTO Key_Sanction (SANID, NAME, ADDMINUS, POINT) "
                tSql &= " VALUES(@SANID, @NAME, @ADDMINUS, @POINT) "
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)
            Case "Key_GradState"
                'KeyType.SelectedValue
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE GradID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                'KeyType.SelectedValue
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr(0) = keycode.Text
                sqldr(1) = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_JoblessWeek"
                'KeyType.SelectedValue
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE JoblessID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                'KeyType.SelectedValue
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr(0) = keycode.Text
                sqldr(1) = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_HourRan"
                'KeyType.SelectedValue
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE HRID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                'KeyType.SelectedValue
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr(0) = keycode.Text
                sqldr(1) = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_TrainExp"
                'KeyType.SelectedValue 
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE TEID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                'KeyType.SelectedValue
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr(0) = keycode.Text
                sqldr(1) = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_HandicatType"
                'KeyType.SelectedValue 
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE HandTypeID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                'KeyType.SelectedValue
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr(0) = keycode.Text
                sqldr(1) = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_HandicatLevel"
                'KeyType.SelectedValue 
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE HandLevelID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                'KeyType.SelectedValue
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr(0) = keycode.Text
                sqldr(1) = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_HandicatType2"
                'KeyType.SelectedValue 
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE HandTypeID2 = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                'KeyType.SelectedValue
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr(0) = keycode.Text
                sqldr(1) = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_HandicatLevel2"
                'KeyType.SelectedValue
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE HandLevelID2 = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                'KeyType.SelectedValue
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr(0) = keycode.Text
                sqldr(1) = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_Exam"
                'KeyType.SelectedValue
                StrSql = "SELECT ExamID,Name FROM " & v_KeyType & " WHERE ExamID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                'KeyType.SelectedValue
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("ExamID") = keycode.Text
                sqldr("Name") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_CostItem"
                'KeyType.SelectedValue 
                StrSql = "SELECT * FROM " & v_KeyType & " Where CostID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                'KeyType.SelectedValue
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("CostID") = keycode.Text
                sqldr("CostName") = keyname.Text
                sqldr("Sort") = Me.Sort.Text
                sqldr("ItemageName") = If(txtItemageName.Text <> "", txtItemageName.Text, Convert.DBNull)
                sqldr("ItemCostName") = If(txtItemCostName.Text <> "", txtItemCostName.Text, Convert.DBNull)
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_CostItem2"
                'KeyType.SelectedValue 
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE CostID='" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                'KeyType.SelectedValue
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("CostID") = keycode.Text
                sqldr("CostName") = keyname.Text
                sqldr("Sort") = Me.Sort.Text
                sqldr("ItemCostName") = If(txtItemCostName.Text <> "", txtItemCostName.Text, Convert.DBNull)
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_Budget"
                StrSql = "SELECT * FROM Key_Budget WHERE BudID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                'sqldr = DbAccess.GetInsertRow("Key_Budget", sqlTable, sqlAdapter)
                'sqldr("BudID") = keycode.Text
                'sqldr("BudName") = keyname.Text
                'sqlTable.Rows.Add(sqldr)
                'sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)"

                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("BUDID", keycode.Text)
                parms.Add("BUDNAME", keyname.Text)
                Dim tSql As String = "INSERT INTO Key_Budget (BUDID, BUDNAME) "
                tSql &= " VALUES(@BUDID, @BUDNAME) "
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)
            Case "Key_DGTHour"
                StrSql = "SELECT * FROM Key_DGTHour WHERE DGID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                'sqldr = DbAccess.GetInsertRow("Key_DGTHour", sqlTable, sqlAdapter)
                'sqldr("DGID") = keycode.Text
                'sqldr("DGName") = keyname.Text
                'sqldr("DGHour") = Me.DGHour.Text
                'sqlTable.Rows.Add(sqldr)
                'sqlAdapter.Update(sqlTable)
                '(新增一筆資料到Table裡)"

                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("DGID", keycode.Text)
                parms.Add("DGNAME", keyname.Text)
                parms.Add("DGHOUR", DGHour.Text)
                Dim tSql As String = "INSERT INTO Key_DGTHour (DGID, DGNAME, DGHOUR) "
                tSql &= " VALUES(@DGID, @DGNAME, @DGHOUR) "
                DbAccess.ExecuteNonQuery(tSql, objconn, parms)

            Case "Key_Salary"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE SalID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("SalID") = keycode.Text
                sqldr("SalName") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_GetJob"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE GJID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("GJID") = keycode.Text
                sqldr("GJName") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)

            Case "Key_NonGetJob"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE NGJID='" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("NGJID") = keycode.Text
                sqldr("NGJName") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_BusVisitCase"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE BVCID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("BVCID") = keycode.Text
                sqldr("BVCName") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_AgeRange"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE ARID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("ARID") = keycode.Text
                sqldr("ARName") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_WorkYear"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE WYID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("WYID") = keycode.Text
                sqldr("WYName") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_Emp"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE KEID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("KEID") = keycode.Text
                sqldr("KEName") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_KindofJob"
            Case "Key_ProSkill"
            Case "Key_Trade"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE TradeID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("TradeID") = keycode.Text
                sqldr("TradeName") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_NotOpenReason"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE NORID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("NORID") = keycode.Text
                sqldr("NORName") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_ClassCatelog"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE CCID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("CCID") = keycode.Text
                sqldr("CCName") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_CheckRan"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE CHID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("CHID") = keycode.Text
                sqldr("CheckRanName") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "ID_Invest"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE IVID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("IVID") = keycode.Text
                sqldr("InvestName") = keyname.Text
                sqldr("ModifyAcct") = sm.UserInfo.UserID
                sqldr("ModifyDate") = Now
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "OB_Funds"
                'StrSql = "Select FID as KID, FName as Name From " & v_KeyType 'KeyType.SelectedValue
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE FID = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr("FID") = keycode.Text
                sqldr("FName") = keyname.Text
                sqldr("ModifyAcct") = sm.UserInfo.UserID
                sqldr("ModifyDate") = Now
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_Native"
                Const Cst_keycode As String = "KNID"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE " & Cst_keycode & " = '" & keycode.Text & "'"
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr(Cst_keycode) = keycode.Text
                sqldr("Name") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_SurveyType"
                Const Cst_keycode As String = "STID"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE " & Cst_keycode & " = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr(Cst_keycode) = keycode.Text
                sqldr("STName") = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_JobGroup" '職群鍵詞檔
                Const Cst_keycode As String = "JGID"
                Const Cst_keyname As String = "JGNAME"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE " & Cst_keycode & " = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr(Cst_keycode) = keycode.Text
                sqldr(Cst_keyname) = keyname.Text
                Sort.Text = TIMS.ClearSQM(Sort.Text)
                sqldr("Sort") = If(Me.Sort.Text <> "", CInt(Val(Me.Sort.Text)), Convert.DBNull) ' Sort.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_EnterPoint" '錄訓百分比代碼
                Const Cst_keycode As String = "KID"
                Const Cst_keyname As String = "KNAME"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE " & Cst_keycode & " = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                sqldr(Cst_keycode) = keycode.Text
                sqldr(Cst_keyname) = keyname.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case "Key_TableMgr"
                StrSql = "SELECT * FROM " & v_KeyType & " WHERE SerNum = '" & keycode.Text & "' "
                If DbAccess.GetCount(StrSql, objconn) > 0 Then
                    Common.MessageBox(Page, "此鍵值代碼已存在" + keycode.Text)
                    Exit Sub
                End If
                sqldr = DbAccess.GetInsertRow(v_KeyType, sqlTable, sqlAdapter, objconn)
                'sqldr("SerNum") = keycode.Text '自動新增代碼
                'KEY_TABLEMGR_SERNUM_SEQ
                sqldr("SerNum") = DbAccess.GetNewId(objconn, "KEY_TABLEMGR_SERNUM_SEQ,KEY_TABLEMGR,SERNUM")
                sqldr("KeyType") = keyname.Text
                sqldr("KeyTable") = txtKeyTable.Text
                sqlTable.Rows.Add(sqldr)
                sqlAdapter.Update(sqlTable)
            Case Else
                Common.MessageBox(Page, "尚無修改程序，請洽系統管理者!!")
                Exit Sub
        End Select

    End Sub

    ''' <summary>'新增</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CmdAppend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAppend.Click
        Call Insert_Key_Table()
        'btnSearch_Click(sender, e)
        SbSearch1()
    End Sub

    ''' <summary> 系統參數重新載入 dirConfig </summary>
    Sub Utl_ResetDirConfig()
        Call TIMS.Utl_SetConfigVAL(objconn)
        'If TIMS.dirConfig Is Nothing Then Return
        'TIMS.dirConfig.Clear()
        'TIMS.dirConfig = Nothing
    End Sub

    ''' <summary>
    ''' 系統參數重新載入
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnConfigReset_Click(sender As Object, e As EventArgs) Handles btnConfigReset.Click
        Call Utl_ResetDirConfig()
    End Sub

    ''' <summary>更新計畫問卷</summary>
    ''' <param name="HHt"></param>
    ''' <param name="oConn"></param>
    Public Shared Sub UPD_PLAN_QUESTIONARY(ByRef HHt As Hashtable, ByRef oConn As SqlConnection)
        Dim v_keycode As String = TIMS.GetMyValue2(HHt, "keycode")
        Dim v_QuesType As String = TIMS.GetMyValue2(HHt, "QuesType")
        Dim dt As New DataTable
        'TIMS.OpenDbConn(oConn)
        Dim s_sql As String = "SELECT TPlanID,QID FROM PLAN_QUESTIONARY WHERE TPlanID=@TPLANID "
        Dim sCmd As New SqlCommand(s_sql, oConn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("TPLANID", SqlDbType.VarChar).Value = v_keycode 'keycode.Text
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            Dim u_sql As String = " UPDATE PLAN_QUESTIONARY SET QID=@QID WHERE TPLANID=@TPLANID" & vbCrLf
            Dim uCmd As New SqlCommand(u_sql, oConn)
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("QID", SqlDbType.Int).Value = Val(If(v_QuesType <> "", v_QuesType, Cst_defaultQID))
                .Parameters.Add("TPLANID", SqlDbType.VarChar).Value = v_keycode 'keycode.Text
                .ExecuteNonQuery()
            End With
        Else
            Dim i_sql As String = " INSERT INTO PLAN_QUESTIONARY(TPLANID,QID) VALUES(@TPLANID,@QID)" & vbCrLf
            Dim iCmd As New SqlCommand(i_sql, oConn)
            With iCmd
                .Parameters.Clear()
                .Parameters.Add("TPLANID", SqlDbType.VarChar).Value = v_keycode ' keycode.Text
                .Parameters.Add("QID", SqlDbType.Int).Value = Val(If(v_QuesType <> "", v_QuesType, Cst_defaultQID))
                .ExecuteNonQuery()
            End With
        End If
    End Sub

End Class