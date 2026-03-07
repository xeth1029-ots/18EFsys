Partial Class TC_01_005_import
    Inherits AuthBasePage

    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False

    Dim FF3 As String = ""
    Const Cst_課程代碼 As Integer = 0
    Const Cst_課程名稱 As Integer = 1
    Const Cst_小時數 As Integer = 2
    Const Cst_學術科 As Integer = 3
    Const Cst_一般專業 As Integer = 4
    Const Cst_主課程代碼 As Integer = 5
    Const Cst_計畫年度 As Integer = 6

    Const Cst_歸屬班別代碼 As Integer = 7
    Const Cst_行業別代碼 As Integer = 8
    Const Cst_訓練職類 As Integer = 9
    Const Cst_是否有效 As Integer = 10
    Const Cst_AllCellsLength As Integer = 11

    'Protected WithEvents PageControler1 As PageControler
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
            Table4.Visible = False '錯誤訊息表
            'CSVPanel.Visible = False 'CVS匯入功能
            'Button1.Visible = False 'CVS匯入功能鈕
        End If

        If Not Session("MySreach") Is Nothing Then
            Me.ViewState("MySreach") = Session("MySreach")
            Session("MySreach") = Nothing
        End If
    End Sub

    Function WrongTable(ByVal drArray As Array, ByVal Reason As String, ByVal dt As DataTable, ByVal Index As Integer) As DataTable
        'Dim dr As DataRow
        'Dim i As Integer

        Dim dr As DataRow = dt.NewRow
        dt.Rows.Add(dr)

        dr("serial") = Index
        If drArray.Length = Cst_AllCellsLength Then
            dr("CourseID") = Left(drArray(Cst_課程代碼).ToString, 10)
            dr("CourseName") = Left(drArray(Cst_課程名稱).ToString, 50)
            dr("Hours") = drArray(Cst_小時數)
            dr("Classification1") = drArray(Cst_學術科)
            dr("Classification2") = drArray(Cst_一般專業)
            dr("MainCourID") = drArray(Cst_主課程代碼)

            dr("Years") = drArray(Cst_計畫年度)
            dr("CLSID") = drArray(Cst_歸屬班別代碼)
            dr("BusID") = drArray(Cst_行業別代碼)
            dr("TMID") = drArray(Cst_訓練職類)
            dr("Valid") = drArray(Cst_是否有效)
        Else
            If drArray.Length < Cst_AllCellsLength Then
                For i As Integer = 0 To drArray.Length - 1
                    dr(i + 1) = drArray(i)
                Next
            ElseIf drArray.Length > Cst_AllCellsLength Then
                For i As Integer = 0 To Cst_AllCellsLength - 1
                    dr(i + 1) = drArray(i)
                Next
            End If
        End If
        dr("reason") = Reason
        Return dt
    End Function

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim drv As DataRowView = e.Item.DataItem
            e.Item.Cells(1).Attributes("onmouseover") = "show_reason('data" & e.Item.ItemIndex & "')"
            e.Item.Cells(1).Attributes("onmouseout") = "dis_reason('data" & e.Item.ItemIndex & "')"
            e.Item.Cells(1).Style.Item("CURSOR") = "hand"
            e.Item.Cells(1).Text += "<br><div align=""right""><table class=""font"" id=""data" & e.Item.ItemIndex & """ style=""position:absolute;display=none"" width=""500"" bgcolor=""#DDE2FB"" bordercolor=""#BBC6F7"" cellspacing=""0"" border=""1"">"
            e.Item.Cells(1).Text += "<tr>"
            e.Item.Cells(1).Text += "<td>"
            e.Item.Cells(1).Text += drv("reason")
            e.Item.Cells(1).Text += "</td>"
            e.Item.Cells(1).Text += "</tr>"
            e.Item.Cells(1).Text += "</table></div>"
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Session("MySreach") = Me.ViewState("MySreach")
        'Response.Redirect("TC_01_005.aspx?ID=" & Request("ID"))
        Dim url1 As String = "TC_01_005.aspx?ID=" & Request("ID")
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Sub sImportFile2(ByRef FullFileName2 As String)
        '上傳檔案
        File2.PostedFile.SaveAs(FullFileName2)

        Dim dt_xls As DataTable = Nothing
        Dim Reason As String = "" '儲存錯誤的原因
        '取得內容
        If (flag_File1_xls) Then
            dt_xls = TIMS.GetDataTable_XlsFile(FullFileName2, "", Reason, "課程代碼", "課程名稱")
            If Reason <> "" Then
                Common.MessageBox(Me, "無法匯入!!" & Reason)
                Exit Sub
            End If
        End If
        If (flag_File1_ods) Then dt_xls = TIMS.GetDataTable_ODSFile(FullFileName2)
        'If (flag_File1_csv) Then dt_xls = TIMS.GetDataTable_CSVFile(FullFileName2)

        '刪除檔案
        'IO.File.Delete(FullFileName1)
        TIMS.MyFileDelete(FullFileName2)

        If dt_xls.Rows.Count = 0 Then
            Common.MessageBox(Me, "資料有誤，故無法匯入，請修正匯入檔案，謝謝")
            Exit Sub
        End If

        '建立錯誤資料格式Table----------------Start
        Dim dtTemp As New DataTable
        dtTemp.Columns.Add(New DataColumn("serial"))
        dtTemp.Columns.Add(New DataColumn("CourseID"))
        dtTemp.Columns.Add(New DataColumn("CourseName"))
        dtTemp.Columns.Add(New DataColumn("Hours"))
        dtTemp.Columns.Add(New DataColumn("Classification1"))
        dtTemp.Columns.Add(New DataColumn("Classification2"))
        dtTemp.Columns.Add(New DataColumn("MainCourID"))

        dtTemp.Columns.Add(New DataColumn("Years"))
        dtTemp.Columns.Add(New DataColumn("CLSID"))
        dtTemp.Columns.Add(New DataColumn("Valid"))
        dtTemp.Columns.Add(New DataColumn("BusID"))
        dtTemp.Columns.Add(New DataColumn("TMID"))
        dtTemp.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table----------------End

        Dim aryCNum As New ArrayList '避免重複暫存
        Dim sql As String = ""
        'xls 方式 讀取寫入資料庫

        Dim colArray As Array

        'select BusID,TrainID,count(1),max(tmid),min(tmid)
        'FROM VIEW_TRAINTYPE
        'where 1=1
        'AND BUSID NOT IN ('G','H')
        'and BusID is not null
        'and TrainID is not null
        'group by BusID,TrainID
        'having count(1)>1
        'order by 1 ,2
        sql = "SELECT * FROM VIEW_TRAINTYPE WHERE 1=1 AND BUSID NOT IN ('G','H')"
        Dim dtTMID As DataTable = DbAccess.GetDataTable(sql, objconn)

        sql = ""
        sql &= " SELECT CLSID,CLASSID,CLASSNAME,CLASSENAME,TPLANID" & vbCrLf
        'sql += " dbo.SUBSTR(CONTENT, 1, 4000) CONTENT,  " & vbCrLf
        sql &= " ,CONTENT" & vbCrLf
        sql &= " ,TMID, DistID, MODIFYACCT, MODIFYDATE, CJOB_UNKEY, Years " & vbCrLf
        sql &= " FROM ID_Class" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND DistID='" & sm.UserInfo.DistID & "'" & vbCrLf '該轄區
        sql &= " AND TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf '該計畫
        Dim dtCLSID As DataTable = DbAccess.GetDataTable(sql, objconn)

        '建立資料表，以機構、RID 查詢
        Dim da As SqlDataAdapter = Nothing
        sql = "SELECT * FROM COURSE_COURSEINFO WHERE OrgID='" & sm.UserInfo.OrgID & "' AND RID='" & sm.UserInfo.RID & "'"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, da, objconn)

        Dim iRowIndex As Integer = 1
        Reason = ""
        For i As Integer = 0 To dt_xls.Rows.Count - 1
            If iRowIndex <> 0 Then
                colArray = dt_xls.Rows(i).ItemArray
                Reason = checkdata2(colArray, aryCNum, dt, dtTMID, dtCLSID)

                If colArray(Cst_主課程代碼).ToString = "" Then
                    If Reason = "" Then
                        Dim dr As DataRow = dt.NewRow
                        dt.Rows.Add(dr)
                        Dim iCOURID As Integer = DbAccess.GetNewId(objconn, "COURSE_COURSEINFO_COURID_SEQ,COURSE_COURSEINFO,COURID")
                        dr("COURID") = iCOURID
                        dr("RID") = sm.UserInfo.RID
                        dr("OrgID") = sm.UserInfo.OrgID
                        dr("CourseID") = Left(colArray(Cst_課程代碼).ToString, 10)
                        aryCNum.Add(dr("CourseID"))
                        dr("CourseName") = Left(colArray(Cst_課程名稱).ToString, 50)
                        dr("Hours") = If(colArray(Cst_小時數).ToString <> "", colArray(Cst_小時數), Convert.DBNull)
                        dr("Classification1") = colArray(Cst_學術科)
                        dr("Classification2") = colArray(Cst_一般專業)
                        dr("MainCourID") = Convert.DBNull 'Cst_主課程代碼
                        '歸屬班別代碼
                        Dim v_CLSID As String = ""
                        If colArray(Cst_歸屬班別代碼).ToString <> "" AndAlso colArray(Cst_計畫年度).ToString <> "" Then
                            v_CLSID = getMainCLSID(dtCLSID, sm.UserInfo.DistID, sm.UserInfo.TPlanID, colArray(Cst_歸屬班別代碼).ToString, colArray(Cst_計畫年度).ToString)
                        End If
                        dr("CLSID") = If(v_CLSID <> "", v_CLSID, Convert.DBNull)

                        Dim v_TMID As String = ""
                        If colArray(Cst_行業別代碼).ToString <> "" And colArray(Cst_訓練職類).ToString <> "" Then
                            FF3 = "BusID='" & colArray(Cst_行業別代碼).ToString & "' and TrainID='" & colArray(Cst_訓練職類).ToString & "'"
                            If dtTMID.Select(FF3).Length > 0 Then v_TMID = dtTMID.Select(FF3)(0)("TMID")
                        End If
                        dr("TMID") = If(v_TMID <> "", v_TMID, Convert.DBNull)
                        dr("Valid") = If(colArray(Cst_是否有效).ToString.ToUpper = "Y", "Y", "N")
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                    Else
                        dtTemp = WrongTable(colArray, Reason, dtTemp, iRowIndex)
                    End If
                End If
            End If
            iRowIndex += 1
        Next
        DbAccess.UpdateDataTable(dt, da)

        '找出有沒有主課程代碼的課程，填入流水號-----   Start
        sql = "SELECT * FROM COURSE_COURSEINFO WHERE RID='" & sm.UserInfo.RID & "'"
        dt = DbAccess.GetDataTable(sql, da, objconn)

        Reason = ""
        iRowIndex = 1
        For i As Integer = 0 To dt_xls.Rows.Count - 1
            If iRowIndex <> 0 Then
                colArray = dt_xls.Rows(i).ItemArray
                Reason = checkdata2(colArray, aryCNum, dt, dtTMID, dtCLSID)

                If colArray(Cst_主課程代碼).ToString <> "" Then
                    If Reason = "" Then
                        Dim dr As DataRow = dt.NewRow
                        dt.Rows.Add(dr)
                        Dim iCOURID As Integer = DbAccess.GetNewId(objconn, "COURSE_COURSEINFO_COURID_SEQ,COURSE_COURSEINFO,COURID")
                        dr("COURID") = iCOURID
                        dr("RID") = sm.UserInfo.RID
                        dr("OrgID") = sm.UserInfo.OrgID
                        dr("CourseID") = Left(colArray(Cst_課程代碼).ToString, 10)
                        aryCNum.Add(dr("CourseID"))
                        dr("CourseName") = Left(colArray(Cst_課程名稱).ToString, 50)
                        dr("Hours") = If(colArray(Cst_小時數).ToString <> "", colArray(Cst_小時數), Convert.DBNull)
                        dr("Classification1") = colArray(Cst_學術科)
                        dr("Classification2") = colArray(Cst_一般專業)
                        'Cst_主課程代碼
                        FF3 = "CourseID='" & colArray(Cst_主課程代碼) & "'"
                        Dim v_MainCourID As String = ""
                        If dt.Select(FF3).Length > 0 Then v_MainCourID = dt.Select(FF3)(0)("CourID")
                        dr("MainCourID") = If(v_MainCourID <> "", v_MainCourID, Convert.DBNull)
                        '歸屬班別代碼
                        Dim v_CLSID As String = ""
                        If colArray(Cst_歸屬班別代碼).ToString <> "" AndAlso colArray(Cst_計畫年度).ToString <> "" Then
                            v_CLSID = getMainCLSID(dtCLSID, sm.UserInfo.DistID, sm.UserInfo.TPlanID, colArray(Cst_歸屬班別代碼).ToString, colArray(Cst_計畫年度).ToString)
                        End If
                        dr("CLSID") = If(v_CLSID <> "", v_CLSID, Convert.DBNull)
                        Dim v_TMID As String = ""
                        If colArray(Cst_行業別代碼).ToString <> "" And colArray(Cst_訓練職類).ToString <> "" Then
                            FF3 = "BusID='" & colArray(Cst_行業別代碼).ToString & "' and TrainID='" & colArray(Cst_訓練職類).ToString & "'"
                            If dt.Select(FF3).Length > 0 Then v_TMID = dtTMID.Select(FF3)(0)("TMID")
                        End If
                        dr("TMID") = If(v_TMID <> "", v_TMID, Convert.DBNull)
                        dr("Valid") = If(colArray(Cst_是否有效).ToString.ToUpper = "Y", "Y", "N")
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                    Else
                        dtTemp = WrongTable(colArray, Reason, dtTemp, iRowIndex)
                    End If
                End If
            End If
            iRowIndex += 1
        Next
        '找出有沒有主課程代碼的課程，填入流水號-----   End

        DbAccess.UpdateDataTable(dt, da)


        '判斷匯出資料是否有誤
        'Dim explain, explain2 As String
        Dim explain As String = ""
        explain = ""
        explain += "匯入資料共" & dt_xls.Rows.Count & "筆" & vbCrLf
        explain += "成功：" & (dt_xls.Rows.Count - dtTemp.Rows.Count) & "筆" & vbCrLf
        explain += "失敗：" & dtTemp.Rows.Count & "筆" & vbCrLf

        'explain2 = ""
        'explain2 += "匯入資料共" & dt_xls.Rows.Count & "筆\n"
        'explain2 += "成功：" & (dt_xls.Rows.Count - dtTemp.Rows.Count) & "筆\n"
        'explain2 += "失敗：" & dtTemp.Rows.Count & "筆\n"

        Common.MessageBox(Me, explain)
        Table4.Visible = False

        If dtTemp.Rows.Count > 0 Then
            'Common.MessageBox(Me, explain)
            Table4.Visible = True

            Label1.Text = dtTemp.Rows.Count
            PageControler1.DataTableCreate(dtTemp)
        End If

    End Sub

    Private Sub Btn_XlsImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_XlsImport.Click
        Dim sMyFileName As String = ""
        'Dim flag_File1_xls As Boolean = False
        'Dim flag_File1_ods As Boolean = False
        Dim sErrMsg As String = TIMS.ChkFile1(File2, sMyFileName, flag_File1_xls, flag_File1_ods)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Return
        End If
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If flag_File1_xls Then
            If Not TIMS.HttpCHKFile(Me, File2, MyPostedFile, "xls") Then Return
        ElseIf flag_File1_ods Then
            If Not TIMS.HttpCHKFile(Me, File2, MyPostedFile, "ods") Then Return
        End If

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File2.PostedFile.FileName).ToLower()
        sMyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Const Cst_FileSavePath As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)
        Dim FullFileName2 As String = Server.MapPath(Cst_FileSavePath & sMyFileName)
        Call sImportFile2(FullFileName2)
    End Sub

    Function getMainCLSID(ByRef dtCLSID As DataTable,
          ByVal sDistID As String, ByVal sTPlanID As String,
          ByVal sClassID As String, ByVal sYears As String) As String

        Dim Rst As String = "" 'CLSID
        Dim sql As String = ""
        If dtCLSID Is Nothing Then
            sql = ""
            sql &= " SELECT CLSID,CLASSID,CLASSNAME,CLASSENAME,TPLANID " & vbCrLf
            'sql += " ,dbo.SUBSTR(CONTENT, 1, 4000) CONTENT " & vbCrLf
            sql &= " ,CONTENT " & vbCrLf
            sql &= " ,TMID, DistID, MODIFYACCT, MODIFYDATE, CJOB_UNKEY, Years " & vbCrLf
            sql &= " FROM ID_CLASS" & vbCrLf
            sql &= " WHERE 1=1" & vbCrLf
            sql &= " AND DistID='" & sDistID & "'" & vbCrLf '該轄區
            sql &= " AND TPlanID='" & sTPlanID & "'" & vbCrLf '該計畫
            dtCLSID = DbAccess.GetDataTable(sql, objconn)
        End If
        If sClassID <> "" AndAlso sYears <> "" Then
            Dim StrTempSearch As String = "" '搜尋字串組合
            StrTempSearch = "ClassID='" & sClassID & "' "
            StrTempSearch += " AND Years='" & sYears & "' "
            If dtCLSID.Select(StrTempSearch).Length = 1 Then
                Rst = dtCLSID.Select(StrTempSearch)(0)("CLSID")
            End If
        End If

        Return Rst
    End Function

    Function checkdata2(ByVal colArray As Array, ByRef aryCNum As ArrayList, ByVal dt As DataTable,
        ByVal dtTMID As DataTable, ByVal dtCLSID As DataTable) As String

        Dim Reason As String = ""
        'Dim i As Integer

        colArray = TIMS.ChangeColArray(colArray) '格式化null成空白文字

        '表示欄位數量正確
        If colArray.Length = Cst_AllCellsLength Then
            '檢查之前是否有輸入相同的課程代碼

            If Convert.ToString(colArray(Cst_課程代碼)) <> "" Then colArray(Cst_課程代碼) = Trim(Convert.ToString(colArray(Cst_課程代碼)))
            If Convert.ToString(colArray(Cst_課程代碼)) = "" Then
                Reason += "課程代碼不可為空<BR>"
            End If

            If Convert.ToString(colArray(Cst_課程名稱)) <> "" Then colArray(Cst_課程名稱) = Trim(Convert.ToString(colArray(Cst_課程名稱)))
            If Convert.ToString(colArray(Cst_課程名稱)) = "" Then
                Reason += "課程名稱不可為空<BR>"
            End If

            For i As Integer = 0 To aryCNum.Count - 1
                If aryCNum(i) = Convert.ToString(colArray(Cst_課程代碼)) Then
                    Reason += "檔案中有相同的課程代碼!<BR>"
                    Exit For
                End If
            Next

            '檢查資料庫內是否有相同的課程代碼
            If dt.Select("CourseID='" & colArray(Cst_課程代碼) & "'").Length <> 0 Then
                '表示有，不寫入資料庫
                Reason += "課程代碼重複<BR>"
            End If
            If colArray(Cst_課程代碼).ToString.Length > 8 Then
                Reason += "課程代碼不能超過8碼<BR>"
            End If

            '檢查所有的欄位是否符合要求
            If colArray(Cst_小時數).ToString <> "" Then
                If Not IsNumeric(colArray(Cst_小時數)) Then
                    Reason += "小時數不為數字<BR>"
                Else
                    If Convert.ToString(CInt(colArray(Cst_小時數))) <> Convert.ToString(colArray(Cst_小時數)) Then
                        Reason += "小時數不可為浮點數<BR>"
                    End If
                End If
            End If

            If IsNumeric(colArray(Cst_學術科)) Then
                If colArray(Cst_學術科) <> 1 And colArray(Cst_學術科) <> 2 Then
                    Reason += "學科[1]術科[2]，值超出範圍<BR>"
                End If
                Select Case colArray(Cst_學術科)
                    Case "1", "2"
                    Case Else
                        Reason += "學科[1]術科[2]，值超出範圍<BR>"
                End Select
            Else
                Reason += "學/術科必須為數字<BR>"
            End If

            Dim tClassification1 As String = Convert.ToString(colArray(Cst_學術科))
            Dim tClassification2 As String = Convert.ToString(colArray(Cst_一般專業))

            If IsNumeric(tClassification2) Then
                Select Case tClassification2
                    Case "0", "1", "2"
                        If tClassification1 = "2" AndAlso tClassification2 <> "2" Then
                            Reason += "術科[2]課程必須為 專業[2]課程<BR>"
                        End If
                    Case Else
                        Reason += "共同[0]一般[1]專業[2]，值超出範圍<BR>"
                End Select
            Else
                Reason += "共同/一般/專業必須為數字<BR>"
            End If

            If colArray(Cst_主課程代碼).ToString <> "" Then
                If dt.Select("CourseID='" & colArray(Cst_主課程代碼).ToString & "'").Length = 0 Then
                    Reason += "查無此主課程代碼<BR>"
                End If
            End If

            If colArray(Cst_計畫年度).ToString <> "" Then
                If Not IsNumeric(colArray(Cst_計畫年度)) Then
                    Reason += "計畫年度必須為數字 如：2010<BR>"
                End If
            Else
                Reason += "計畫年度不可為空<BR>"
            End If

            If Reason = "" Then
                '歸屬班別代碼 check 唯等於1才可繼續
                If colArray(Cst_歸屬班別代碼).ToString <> "" Then
                    Dim StrTempSearch As String = "" '搜尋字串組合
                    StrTempSearch = "ClassID='" & colArray(Cst_歸屬班別代碼).ToString & "' "
                    StrTempSearch += " AND Years='" & colArray(Cst_計畫年度).ToString & "' "
                    'StrTempSearch += " AND TPlanID='" & sm.UserInfo.TPlanID.ToString & "' "

                    If dtCLSID.Select(StrTempSearch).Length = 0 Then
                        Reason += "該計畫年度,查無此歸屬班別代碼<BR>"
                    End If
                    If dtCLSID.Select(StrTempSearch).Length > 1 Then
                        Reason += "該大計畫、計畫年度,搜尋歸屬班別代碼數量有誤大於1，請修正該班別代碼<BR>"
                    End If
                End If
            End If

            If colArray(Cst_行業別代碼).ToString = "" And colArray(Cst_一般專業).ToString <> "0" Then
                Reason += "必須要有行業別代碼<BR>"
            End If
            If colArray(Cst_訓練職類).ToString = "" And colArray(Cst_一般專業).ToString <> "0" Then
                Reason += "必須要有職類代碼<BR>"
            Else
                If colArray(Cst_訓練職類).ToString <> "" Then
                    If IsNumeric(colArray(Cst_訓練職類)) Then
                        If colArray(Cst_訓練職類).ToString.Length <> 4 Then
                            Reason += "職類代碼必須為4位數字<BR>"
                        End If
                    Else
                        Reason += "職類代碼必須為數字<BR>"
                    End If
                End If
            End If
            If colArray(Cst_行業別代碼).ToString <> "" And colArray(Cst_訓練職類).ToString <> "" Then
                If dtTMID.Select("TrainID='" & colArray(Cst_訓練職類).ToString & "' and BusID='" & colArray(Cst_行業別代碼).ToString & "'").Length = 0 Then
                    Reason += "找不到此職類代碼!請檢查行業別或職類代碼是否有誤!"
                End If
            End If
            If colArray(Cst_是否有效).ToString <> "Y" And colArray(9).ToString <> "N" Then
                Reason += "是否有效必須為Y或N值<BR>"
            End If
        Else
            Reason += "檔案欄位數量有誤!<BR>"
        End If

        Return Reason
    End Function

End Class
