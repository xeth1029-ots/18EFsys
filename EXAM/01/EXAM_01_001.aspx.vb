Partial Class EXAM_01_001
    Inherits AuthBasePage

    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False

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

        '載入選擇頁
        PageControler1.PageDataGrid = dg_Sch

        If Not IsPostBack Then
            '載入分署(中心)
            Dim sqlstr As String
            sqlstr = "select * from ID_District "
            If sm.UserInfo.DistID <> "000" Then '系統管理者可查全部
                sqlstr += "where DistID = '" & sm.UserInfo.DistID & "'"
            Else
                sqlstr += "where DistID != '000'"
            End If
            Dim dt As DataTable = DbAccess.GetDataTable(sqlstr, objconn)
            ddl_DistID = TIMS.Get_DistID(ddl_DistID, dt)
            ddl_DistID.Items.Remove(ddl_DistID.Items.FindByValue(""))
        End If
    End Sub

    '查詢
    Sub search()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, dg_Sch) '顯示列數不正確

        Dim sql As String = ""
        sql += " select case " & vbCrLf
        sql += "  when p.Name is not null then CONVERT(varchar, p.Name) + '_' + CONVERT(varchar, IE.Name) " & vbCrLf
        sql += "  when IE.Name is not null then CONVERT(varchar, IE.Name) " & vbCrLf
        sql += "  else '不區分'  end as Name" & vbCrLf
        sql += " ,dbo.NVL(P.Name, dbo.NVL(IE.Name,'不區分')) as PName" & vbCrLf
        sql += " ,case when P.Name is null then '不區分' else CONVERT(varchar, IE.Name) end as cName" & vbCrLf
        'sql += " , P.Name as PName,  dbo.NVL(IE.Name,'不區分') as cName " & vbCrLf
        sql += " ,IE.ETID,D.Name as DistName, case IE.Avail when '1' then '是' else '否' end Avail " & vbCrLf
        sql += " ,IE.Parent, IE.Levels " & vbCrLf
        sql += " ,IE.Avail CAvail" & vbCrLf
        sql += " ,P.Avail PAvail" & vbCrLf
        sql += "from ID_ExamType IE " & vbCrLf
        sql += "join ID_District D on D.DistID=IE.DistID " & vbCrLf
        sql += "LEFT JOIN ID_ExamType P ON IE.Parent=P.ETID" & vbCrLf
        sql += "where 1=1 " & vbCrLf
        If Len(ddl_DistID.SelectedValue) > 0 Then
            sql += "and IE.DistID='" & ddl_DistID.SelectedValue & "'" & vbCrLf
        End If
        If Len(txt_pName.Text) > 0 Then
            sql += "and P.Name like '" & txt_pName.Text & "%' " & vbCrLf
        End If
        If Len(txt_cName.Text) > 0 Then
            sql += "and IE.Name like '" & txt_cName.Text & "%' " & vbCrLf
        End If
        sql += " ORDER BY IE.ETID, 1 " & vbCrLf

        msg.Visible = True
        Panel_View.Visible = False
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            msg.Visible = False
            Panel_View.Visible = True
            PageControler1.PageDataTable = dt '.SqlString = sql
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub btn_Sch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Sch.Click
        Panel_Sch.Visible = True
        Panel_View.Visible = False
        Panel_edit.Visible = False
        Panel_Add.Visible = False

        Call search()
    End Sub

    Private Sub dg_Sch_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dg_Sch.ItemCommand
        Dim btnedit As Button = e.Item.FindControl("btn_edit")
        Me.ViewState("ETID") = btnedit.CommandArgument
        Me.ViewState("ETID") = TIMS.ClearSQM(Me.ViewState("ETID"))
        Dim sql As String
        Dim sqlstr As String
        Dim dr As DataRow
        Select Case e.CommandName
            Case "edit" 'EDIT

                Panel_Sch.Visible = False
                Panel_View.Visible = False
                Panel_edit.Visible = True
                'Me.ddleditParent = Exam.Get_ExamTypeParent(ddleditParent, sm.UserInfo.DistID)
                Me.ddleditParent = cls_Exam.Get_ExamTypeParent(ddleditParent, ddl_DistID.SelectedValue)
                ddl_editDistID = TIMS.Get_DistID(ddl_editDistID)

                sql = " select ie.etid,id.distid,ie.name,ie.avail, IE.Parent from ID_ExamType ie join id_district id on id.distid=ie.distid "
                sql += " where ie.etid=" & Me.ViewState("ETID")

                dr = DbAccess.GetOneRow(sql, objconn)
                ddl_editDistID.Enabled = False
                Common.SetListItem(ddl_editDistID, dr("distid"))

                ddleditParent.Enabled = False
                If Convert.ToString(dr("Parent")) <> "" Then
                    ddleditParent.Enabled = True
                    Common.SetListItem(ddleditParent, Convert.ToString(dr("Parent")))
                End If
                txt_editdname.Text = dr("name")
                '載入分署(中心)＆選定DB分署(中心)編號
                rbl_eavail.SelectedValue = dr("avail")

            Case "del"
                Dim ErrMsg As String = ""
                Dim DEL_flag_OK As Boolean = True '可刪除　FALSE '不可刪除
                Dim dt As DataTable = Nothing

                sql = " select ie.etid from id_examtype ie join exam_question eq on ie.etid=eq.etid"
                sql += " where ie.etid=" & Me.ViewState("ETID")
                dt = DbAccess.GetDataTable(sql, objconn)
                If dt.Rows.Count > 0 Then
                    Common.MessageBox(Me, "此類別尚有題庫故不可刪除!")
                    DEL_flag_OK = False 'Exit Sub
                End If

                sql = "select ie.etid from id_examtype ie where ie.Parent=" & Me.ViewState("ETID")
                dt = DbAccess.GetDataTable(sql, objconn)
                If dt.Rows.Count > 0 Then
                    Common.MessageBox(Me, "此類別尚有子類別故不可刪除!")
                    DEL_flag_OK = False 'Exit Sub
                End If

                If DEL_flag_OK Then
                    sqlstr = "delete id_examtype where etid=" & Me.ViewState("ETID")
                    DbAccess.ExecuteNonQuery(sqlstr, objconn)
                    'Page.RegisterStartupScript("del", "<script>alert('刪除成功!');</script>")
                    Call search()
                Else
                    Common.MessageBox(Me, ErrMsg)
                End If
        End Select
    End Sub

    Private Sub dg_Sch_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_Sch.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem ', ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btnedit As Button = e.Item.FindControl("btn_edit")
                Dim btndel As Button = e.Item.FindControl("btn_del")
                'Dim btnViewItem As Button = e.Item.FindControl("btnViewItem")

                e.Item.Cells(0).Text = Convert.ToString(drv("ETID"))  'e.Item.ItemIndex + 1
                If (e.Item.Cells(3).Text) = "&nbsp;" Then
                    e.Item.Cells(3).Text = "共用"
                End If

                If Convert.ToString(drv("PAvail")) <> "" AndAlso Convert.ToString(drv("CAvail")) = "1" Then
                    If Convert.ToString(drv("PAvail")) <> Convert.ToString(drv("CAvail")) Then
                        e.Item.Cells(1).Text = "<font color='Red'>" & Convert.ToString(drv("pName")) & "(異常停用)</font>"
                        TIMS.Tooltip(e.Item.Cells(1), "異常停用!!")
                    Else
                        e.Item.Cells(1).Text = "<font color='Silver'>" & Convert.ToString(drv("pName")) & "</font>"
                    End If
                End If

                btnedit.CommandArgument = drv("ETID")
                btndel.CommandArgument = drv("ETID")
                'btnViewItem.CommandArgument = drv("ETID")
                'If Convert.ToString(drv("Parent")) = "" Then
                '    btnViewItem.Visible = True '判斷無上層可顯示維護子明細
                'Else
                '    btnViewItem.Visible = False
                'End If
                'btndel.Attributes("onclick") = "return confirm('確定要刪除第 " & e.Item.Cells(0).Text & " 筆資料?');"
                btndel.Attributes("onclick") = "return confirm('確定要刪除 " & e.Item.Cells(0).Text & "號 資料?');"
        End Select

    End Sub

    Private Sub btn_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Add.Click
        Panel_Sch.Visible = False
        Panel_View.Visible = False
        Panel_Add.Visible = True
        txt_distid.Text = ddl_DistID.SelectedItem.Text
        hDistIDVal.Value = ddl_DistID.SelectedValue
        txt_dname.Text = ""
        Me.ddlParent = cls_Exam.Get_ExamTypeParent(ddlParent, hDistIDVal.Value)

    End Sub

    '新增儲存
    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        If txt_dname.Text = "" Then
            Common.MessageBox(Me, "請填寫類別名稱!")
            Exit Sub
        End If

        '檢查是否重複
        Dim sql_check As String = ""
        sql_check = "select * from ID_ExamType where 1=1 "
        sql_check += " and DistID='" & ddl_DistID.SelectedValue & "'"
        sql_check += " and name='" & txt_dname.Text & "'"
        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql_check, objconn)
        If dt.Rows.Count > 0 Then
            Common.MessageBox(Me, "類別名稱或訓練計畫重複!")
            Exit Sub
        End If

        Dim sql As String = ""
        sql = "insert into id_examtype(DistID,Name,Avail,Parent,Levels,ModifyAcct,ModifyDate) "
        sql += "values('" & hDistIDVal.Value & "','" & txt_dname.Text & "'," & rbl_avail.SelectedValue

        If Me.ddlParent.SelectedValue <> "" Then
            sql += "," & Me.ddlParent.SelectedValue & ",1"
        Else
            sql += ",NULL,NULL"
        End If
        sql += ",'" & sm.UserInfo.UserID & "',getdate())"
        DbAccess.ExecuteNonQuery(sql, objconn)

        Common.MessageBox(Me, "存檔成功!")

        btn_Sch_Click(sender, e)

    End Sub

    '離開(新增)
    Private Sub btn_exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_exit.Click
        Call btn_Sch_Click(sender, e)
    End Sub

    '修改儲存
    Private Sub btn_editsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_editsave.Click
        Dim ErrMsg As String
        ErrMsg = ""

        Dim sql_check As String
        Dim sql As String
        If txt_editdname.Text = "" Then
            ErrMsg += "請填寫類別名稱!" & vbCrLf
            'Common.MessageBox(Me, "請填寫類別名稱!")
            'Exit Sub
        End If

        If Me.ddleditParent.Enabled Then
            If ddleditParent.SelectedValue = "" Then
                ErrMsg += "請選擇父層類別!" & vbCrLf
            End If
        End If

        If ErrMsg <> "" Then
            Common.MessageBox(Me, ErrMsg)
            Exit Sub
        End If

        '檢查是否重複
        Dim dt As DataTable = Nothing
        sql_check = "select * from id_examtype where 1=1 "
        sql_check += "and distid='" & ddl_editDistID.SelectedValue & "' "
        sql_check += "and name='" & txt_editdname.Text & "' "
        sql_check += "and etid<>" & Me.ViewState("ETID")
        dt = DbAccess.GetDataTable(sql_check, objconn)
        If dt.Rows.Count > 0 Then
            Common.MessageBox(Me, "類別名稱或訓練計畫重複!")
            Exit Sub
        End If

        sql = "update id_examtype set distid='" & ddl_editDistID.SelectedValue
        sql += "',name='" & txt_editdname.Text & "',modifyacct='" & sm.UserInfo.UserID & "',avail=" & rbl_eavail.SelectedValue
        If Me.ddleditParent.SelectedValue <> "" Then
            sql += ",Parent=" & Me.ddleditParent.SelectedValue & ",Levels=1"
        Else
            sql += ",Parent=NULL,Levels=NULL"
        End If
        sql += ",modifydate=getdate()"
        sql += " where etid=" & Me.ViewState("ETID")
        DbAccess.ExecuteNonQuery(sql, objconn)
        Common.MessageBox(Me, "修改成功!")

        btn_Sch_Click(sender, e)
    End Sub

    '離開(修改)
    Private Sub btn_editexit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_editexit.Click

        Call btn_Sch_Click(sender, e)
    End Sub

    '檢查(匯入Excel)資料
    Function CheckImportData(ByVal col_dr As Array) As String
        Dim Reason As String = ""
        Const Cst_Parent As Integer = 0
        'Const Cst_Name As Integer = 1
        'Const Cst_Avail As Integer = 2
        If col_dr.Length < 3 Then
            Reason += "欄位對應有誤<BR>"
        Else
            Dim vPNameWrFlag As Boolean = False
            If Convert.ToString(col_dr(Cst_Parent)) <> "" Then
                Dim vPName As String = ""
                vPName = cls_Exam.Get_ExamTypePName(Convert.ToString(col_dr(Cst_Parent)), sm.UserInfo.DistID, objconn)
                If vPName = "" Then
                    Reason += "查無上層類別<BR>"
                End If
            End If
        End If
        Return Reason
    End Function

    '匯入Excel
    Sub Insert_DataTableXLS(ByVal dt_xls As DataTable)
        'Const Cst_Spage As String = "cp_08_013"
        Const WrongPageUrl As String = "EXAM_01_001_Wrong.aspx"
        Const Cst_Parent As Integer = 0
        Const Cst_Name As Integer = 1
        Const Cst_Avail As Integer = 2

        '建立sql的津貼連線
        'Dim tConn As SqlConnection = New SqlConnection
        'tConn = DbAccess.GetConnection()
        Call TIMS.OpenDbConn(objconn)

        '檢查重覆資料 '類別名稱 '轄區分署(轄區中心)代碼
        Dim strSql As String = "
SELECT COUNT(1) CNT FROM ID_ExamType
where Name=@Name and DistID=@DistID
"
        Dim SEL_COUNT As New SqlCommand(strSql, objconn)

        Dim strSql_into As String = "
INSERT INTO ID_ExamType(DistID ,Name ,Avail,ModifyAcct ,ModifyDate,Parent ,Levels) 
VALUES (@DistID ,@Name ,@Avail,@ModifyAcct ,GETDATE(),@Parent ,@Levels)"
        Dim insert_cmd As New SqlCommand(strSql_into, objconn)

        '開使處理要匯入的資料
        Dim RowIndex As Integer = 1 '改為Excel位置由1開始再加1等於2
        Dim Reason As String = "" '儲存錯誤的原因
        Dim dtWrong As New DataTable '儲存錯誤資料的DataTable
        Dim drWrong As DataRow
        '建立錯誤資料格式Table
        dtWrong.Columns.Add(New DataColumn("Index")) '序號
        dtWrong.Columns.Add(New DataColumn("Reason")) '問題

        For Each dr As DataRow In dt_xls.Rows
            RowIndex += 1 '改為Excel位置由1開始再加1等於2
            Dim colArray As Array = dr.ItemArray
            Reason = CheckImportData(colArray)

            If Reason <> "" Then
                '錯誤資料，填入錯誤資料表
                drWrong = dtWrong.NewRow
                dtWrong.Rows.Add(drWrong)

                drWrong("Index") = RowIndex '改為Excel位置由1開始再加1等於2
                drWrong("Reason") = Reason
            Else
                '匯入資料
                Dim SqlCount As Integer = 0
                SqlCount = 0
                SEL_COUNT.Parameters.Clear()
                SEL_COUNT.Parameters.Add("Name", SqlDbType.NVarChar).Value = Convert.ToString(dr(Cst_Name))
                SEL_COUNT.Parameters.Add("DistID", SqlDbType.VarChar).Value = sm.UserInfo.DistID

                SqlCount = SEL_COUNT.ExecuteScalar()
                If SqlCount > 0 Then
                    '資料重複
                    '錯誤資料，填入錯誤資料表
                    drWrong = dtWrong.NewRow
                    dtWrong.Rows.Add(drWrong)

                    drWrong("Index") = RowIndex
                    drWrong("Reason") = "(" & Convert.ToString(dr(Cst_Name)) & ")資料重複!!"
                Else
                    'insert
                    insert_cmd.Parameters.Clear()
                    insert_cmd.Parameters.Add("DistID", SqlDbType.VarChar).Value = sm.UserInfo.DistID
                    insert_cmd.Parameters.Add("Name", SqlDbType.NVarChar).Value = Convert.ToString(dr(Cst_Name))
                    If Convert.ToString(dr(Cst_Avail)) = "0" Then
                        insert_cmd.Parameters.Add("Avail", SqlDbType.VarChar).Value = Convert.ToString(dr(Cst_Avail))
                    Else
                        insert_cmd.Parameters.Add("Avail", SqlDbType.VarChar).Value = "1"
                    End If
                    insert_cmd.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    Dim vPNameWrFlag As Boolean = False
                    If Convert.ToString(dr(Cst_Parent)) <> "" Then
                        Dim vPName As String = ""
                        vPName = cls_Exam.Get_ExamTypePName(Convert.ToString(dr(Cst_Parent)), sm.UserInfo.DistID, objconn)
                        If vPName <> "" Then
                            insert_cmd.Parameters.Add("Parent", SqlDbType.VarChar).Value = Convert.ToString(dr(Cst_Parent))
                            insert_cmd.Parameters.Add("Levels", SqlDbType.VarChar).Value = "1"
                            vPNameWrFlag = True
                        End If
                    End If
                    If Not vPNameWrFlag Then
                        insert_cmd.Parameters.Add("Parent", SqlDbType.VarChar).Value = Convert.DBNull
                        insert_cmd.Parameters.Add("Levels", SqlDbType.VarChar).Value = Convert.DBNull
                    End If
                    insert_cmd.ExecuteNonQuery()
                End If

            End If
        Next
        'Call TIMS.CloseDbConn(tConn)

        '判斷匯出資料是否有誤
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

        If dtWrong.Rows.Count = 0 Then
            If Reason = "" Then
                Common.MessageBox(Me, explain)
            Else
                Common.MessageBox(Me, explain & Reason)
            End If
        Else
            Session("MyWrongTable") = dtWrong '塞入session 
            Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視原因?')){window.open('" & WrongPageUrl & "','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
        End If
    End Sub

    '匯入Excel(按鈕)
    Private Sub BtnImport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnImport1.Click
        Session("MyWrongTable") = Nothing
        Dim Upload_Path As String = "~/EXAM/01/Temp/"
        Call TIMS.MyCreateDir(Me, Upload_Path)
        Const Cst_Filetype As String = "xls" '匯入檔案類型
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, Cst_Filetype) Then Return

        'Dim MyFile As System.IO.File
        Dim MyFileName As String
        Dim MyFileType As String
        Dim flag As String

        If File1.Value <> "" Then
            If File1.PostedFile.ContentLength = 0 Then
                Common.MessageBox(Me, "檔案位置錯誤!")
                Exit Sub
            Else
                '取出檔案名稱
                MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)

                '取出檔案類型
                If MyFileName.IndexOf(".") = -1 Then

                    Common.MessageBox(Me, "檔案類型錯誤!")
                    Exit Sub
                Else
                    MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
                    If LCase(MyFileType) = LCase(Cst_Filetype) Then
                        flag = ","
                    Else
                        Common.MessageBox(Me, "檔案類型錯誤，必須為" & UCase(Cst_Filetype) & "檔!")
                        Exit Sub
                    End If
                End If
            End If
            '檢查檔案格式與大小----------   End

            Dim dt_xls As DataTable
            Dim Errmag As String = ""
            '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
            Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
            MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
            Dim filePath1 As String = Server.MapPath(Upload_Path & MyFileName)
            '上傳檔案
            File1.PostedFile.SaveAs(filePath1)
            '取得內容
            dt_xls = TIMS.GetDataTable_XlsFile(Server.MapPath(Upload_Path & MyFileName).ToString, "", Errmag, "子類別名稱")
            '刪除檔案'IO.File.Delete(Server.MapPath(Upload_Path & MyFileName)) 
            TIMS.MyFileDelete(filePath1)

            If Errmag <> "" Then
                Common.MessageBox(Me, Errmag)
                Common.MessageBox(Me, "資料有誤，故無法匯入，請修正Excel檔案，謝謝")
                Exit Sub
            End If

            Call Insert_DataTableXLS(dt_xls)
        Else
            '沒有檔案名稱
            Common.MessageBox(Me, "請選擇匯入檔案的路徑!!")
        End If
    End Sub

End Class
