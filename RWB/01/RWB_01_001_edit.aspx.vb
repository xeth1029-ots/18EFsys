Partial Class RWB_01_001_edit
    Inherits AuthBasePage

    Dim objconn As SqlConnection
    Dim i_contentCount As Integer = 4   '定義目前『內容區塊』最大數為4
    Dim mySavePath As String = ConfigurationManager.AppSettings("UploadPath") + "/NEWS/"   '定義圖片、附件欲存放的路徑
    Dim ImageUrl_Path1 As String = "../../Upload/NEWS/"

    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call sCreate1() '頁面初始化
        End If
    End Sub

    ''' <summary>
    ''' 頁面初始化
    ''' </summary>
    Sub sCreate1()
        trC_URL.Visible = False
        trC_CONTENT1.Visible = False

        hid_V_SEQNO.Value = ""
        hid_f_grp.Value = ""

        ddlType.Enabled = True
        Const cst_rwb01001_add As String = "rwb01001_add"
        Dim V_DDLTYPE As String = ""
        If Session(cst_rwb01001_add) IsNot Nothing Then
            Dim strSession As String = Convert.ToString(Session(cst_rwb01001_add))
            V_DDLTYPE = TIMS.GetMyValue(strSession, "ddlType")
            If V_DDLTYPE <> "" Then
                '新增給個預設值
                Select Case V_DDLTYPE
                    Case "011"
                        ddlType = TIMS.Get_RWBFUNTYPE(ddlType, 4) '011 (sp)
                        Common.SetListItem(ddlType, "011")
                        Call SET_Visible_ALL(V_DDLTYPE)

                    Case Else
                        ddlType = TIMS.Get_RWBFUNTYPE(ddlType, 3) '預設 1/2/3
                        Common.SetListItem(ddlType, V_DDLTYPE)

                End Select
            End If
        End If
        Session(cst_rwb01001_add) = Nothing

        txtCDATE1.Text = If(flag_ROC, TIMS.Cdate17(DateTime.Now.ToString("yyyy/MM/dd")), DateTime.Now.ToString("yyyy/MM/dd"))
        ddlC_SDATE_hh1.Items.Clear()
        ddlC_EDATE_hh1.Items.Clear()
        For i As Integer = 0 To 23
            ddlC_SDATE_hh1.Items.Add(New ListItem(i.ToString.PadLeft(2, "0"), i.ToString.PadLeft(2, "0")))
            ddlC_EDATE_hh1.Items.Add(New ListItem(i.ToString.PadLeft(2, "0"), i.ToString.PadLeft(2, "0")))
        Next

        ddlC_SDATE_mm1.Items.Clear()
        ddlC_EDATE_mm1.Items.Clear()
        For j As Integer = 0 To 59
            ddlC_SDATE_mm1.Items.Add(New ListItem(j.ToString.PadLeft(2, "0"), j.ToString.PadLeft(2, "0")))
            ddlC_EDATE_mm1.Items.Add(New ListItem(j.ToString.PadLeft(2, "0"), j.ToString.PadLeft(2, "0")))
        Next

        Common.SetListItem(ddlC_SDATE_hh1, "00")
        Common.SetListItem(ddlC_SDATE_mm1, "00")
        Common.SetListItem(ddlC_EDATE_hh1, "23")
        Common.SetListItem(ddlC_EDATE_mm1, "59")

        '讀取已存在DB的資料內容 Request("A") E:EDIT
        Select Case TIMS.ClearSQM(Request("A"))
            Case "E" '修改
                '『內容區塊』元件初始化
                Dim tDT As DataTable = New DataTable
                tDT = CreateAreaData(i_contentCount)
                If tDT IsNot Nothing Then
                    If tDT.Rows.Count > 0 Then
                        gv1.DataSource = tDT
                        gv1.DataBind()
                    End If
                End If

                Dim rFUNID_E As String = TIMS.DecryptAes(TIMS.ClearSQM(Request("FUNID_E"))) '解
                Dim rSEQNO_E As String = TIMS.DecryptAes(TIMS.ClearSQM(Request("SEQNO_E"))) '解
                Dim rSEQNO As String = TIMS.ClearSQM(Request("SEQNO"))
                If rSEQNO_E <> "" AndAlso rSEQNO_E = rSEQNO Then hid_V_SEQNO.Value = rSEQNO

                If hid_V_SEQNO.Value <> "" Then
                    Dim iSEQNO As Integer = Val(hid_V_SEQNO.Value)
                    Load_TB_CONTENT(iSEQNO, rFUNID_E)
                    Load_TB_CONTENT_SECTION(iSEQNO)
                    'show hid_f_grp.Value
                    Load_TB_FILE()
                End If
                'If hid_V_SEQNO.Value <> "" Then LoadData(Val(hid_V_SEQNO.Value))
            Case Else
                '新增-[其它]
                Select Case V_DDLTYPE
                    Case "011"
                    Case Else
                        '『內容區塊』元件初始化
                        Dim tDT As DataTable = New DataTable
                        tDT = CreateAreaData(i_contentCount)
                        If tDT IsNot Nothing Then
                            If tDT.Rows.Count > 0 Then
                                gv1.DataSource = tDT
                                gv1.DataBind()
                            End If
                        End If
                End Select
        End Select

    End Sub

    Sub SET_Visible_ALL(ByVal S_FUNID As String)
        Select Case S_FUNID
            Case "011"
                ddlType.Enabled = False

                trLINKURL1.Visible = False
                trLINKURL2.Visible = False
                trUPFILE1.Visible = False

                trC_URL.Visible = True
                trC_CONTENT1.Visible = True
        End Select
    End Sub

#Region "自定義虛擬表格資料"

    Function CreateAreaData(ByVal i_myRows As Integer) As DataTable
        'i_contentCount
        Dim sql2 As String = ""
        sql2 &= " SELECT SEC_NO id,'區塊'+convert(varchar,SEC_NO) item" & vbCrLf
        sql2 &= " ,SEQNO, CONTENTID, SEC_NO, SEC_CONTENT, SEC_PICTURE" & vbCrLf
        sql2 &= " ,ALIGN_TYPE, MODIFYACCT, MODIFYDATE, SEC_PICTURE_ALT" & vbCrLf
        sql2 &= " FROM TB_CONTENT_SECTION" & vbCrLf
        sql2 &= " WHERE 1<>1" & vbCrLf
        Dim parms2 As Hashtable = New Hashtable()
        Dim dt2 As DataTable = DbAccess.GetDataTable(sql2, objconn, parms2)
        'Dim myTable As DataTable=New DataTable
        'Dim column As DataColumn
        'Dim row As DataRow
        'column=New DataColumn()
        'column.DataType=System.Type.GetType("System.Int32")
        'column.ColumnName="id"
        'myTable.Columns.Add(column)
        'column=New DataColumn()
        'column.DataType=Type.GetType("System.String")
        'column.ColumnName="item"
        'myTable.Columns.Add(column)
        'Dim i As Integer
        For i As Integer = 0 To i_myRows - 1
            Dim dr1 As DataRow = dt2.NewRow()
            dt2.Rows.Add(dr1)
            dr1("id") = (i + 1)
            dr1("item") = "區塊" & (i + 1).ToString()
        Next
        Return dt2
    End Function

#End Region

    '資料讀取
    'Private Sub LoadData(ByVal iSEQNO As Integer)
    'LoadData1(iSEQNO)
    'LoadData2(iSEQNO)
    'LoadData3()
    'End Sub

    '(1)讀取基本內容區塊
    Private Sub Load_TB_CONTENT(ByVal iSEQNO As Integer, ByVal s_FUNID As String)
        s_FUNID = If(s_FUNID <> "", s_FUNID, "001")
        Dim sql As String = ""
        sql &= " SELECT ROW_NUMBER() OVER(ORDER BY a.C_UDATE DESC) AS ROWNUM" & vbCrLf
        sql &= " ,FORMAT(a.C_SDATE, 'yyyy-MM-dd') CSDATE" & vbCrLf
        sql &= " ,FORMAT(a.C_EDATE, 'yyyy-MM-dd') CEDATE" & vbCrLf
        sql &= " ,FORMAT(a.C_CDATE, 'yyyy-MM-dd') CCDATE" & vbCrLf
        sql &= " ,FORMAT(a.C_SDATE, 'HH') CSDATEHH" & vbCrLf
        sql &= " ,FORMAT(a.C_EDATE, 'HH') CEDATEHH" & vbCrLf
        sql &= " ,FORMAT(a.C_SDATE, 'mm') CSDATEMM" & vbCrLf
        sql &= " ,FORMAT(a.C_EDATE, 'mm') CEDATEMM" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, a.C_SDATE, 111) CSDATED" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, a.C_EDATE, 111) CEDATED" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, a.C_CDATE, 111) CCDATED" & vbCrLf
        sql &= " ,a.SEQNO" & vbCrLf
        sql &= " ,a.FUNID" & vbCrLf
        sql &= " ,a.SUB_FUNID" & vbCrLf
        sql &= " ,a.C_SDATE ,a.C_EDATE" & vbCrLf
        sql &= " ,a.C_TITLE ,a.C_URL" & vbCrLf
        sql &= " ,a.C_LINKURL1,a.C_LINKURL2" & vbCrLf
        sql &= " ,a.C_LINKURL3,a.C_LINKURL4 ,a.C_LINKURL5" & vbCrLf
        sql &= " ,a.C_CONTENT1,a.C_CONTENT2,a.C_CONTENT3" & vbCrLf
        sql &= " ,a.C_CDATE" & vbCrLf
        sql &= " ,a.C_CACCT" & vbCrLf
        sql &= " ,a.C_UDATE" & vbCrLf
        sql &= " ,a.C_UACCT" & vbCrLf
        sql &= " ,a.C_STATUS" & vbCrLf
        sql &= " ,a.F_GROUPID" & vbCrLf
        sql &= " ,a.C_SORT1" & vbCrLf
        sql &= " FROM TB_CONTENT a" & vbCrLf
        sql &= " WHERE a.FUNID =@FUNID" & vbCrLf
        sql &= " AND a.SEQNO=@SEQNO" & vbCrLf

        Dim parms As New Hashtable From {
            {"SEQNO", iSEQNO},
            {"FUNID", s_FUNID}
        }

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        If dt.Rows.Count = 0 Then Exit Sub

        Dim dr As DataRow = dt.Rows(0)
        hid_f_grp.Value = Convert.ToString(dr("F_GROUPID"))

        'ddlType.SelectedValue=Convert.ToString(dr("SUB_FUNID"))
        'Dim s_FUNID As String=Convert.ToString(dr("FUNID"))
        Select Case s_FUNID 'Convert.ToString(dr("FUNID"))
            Case "011"
                ddlType = TIMS.Get_RWBFUNTYPE(ddlType, 4)
                Common.SetListItem(ddlType, Convert.ToString(dr("FUNID")))
                SET_Visible_ALL(s_FUNID)

                txtC_URL.Text = Convert.ToString(dr("C_URL"))
                txtC_CONTENT1.Text = Convert.ToString(dr("C_CONTENT1"))
            Case Else
                ddlType = TIMS.Get_RWBFUNTYPE(ddlType, 3)
                Common.SetListItem(ddlType, Convert.ToString(dr("SUB_FUNID")))
        End Select

        If flag_ROC Then
            txtCDATE1.Text = TIMS.Cdate17(Convert.ToString(dr("CCDATED")))  'edit，by:20181019
            txtSDATE1.Text = TIMS.Cdate17(Convert.ToString(dr("CSDATED")))  'edit，by:20181019
        Else
            txtCDATE1.Text = Convert.ToString(dr("CCDATED"))
            txtSDATE1.Text = Convert.ToString(dr("CSDATED"))
        End If
        Common.SetListItem(ddlC_SDATE_hh1, Convert.ToString(dr("CSDATEHH")))
        Common.SetListItem(ddlC_SDATE_mm1, Convert.ToString(dr("CSDATEMM")))

        txtTitle.Text = Convert.ToString(dr("C_TITLE"))
        txtLINKURL1.Text = Convert.ToString(dr("C_LINKURL1"))
        txtLINKURL2.Text = Convert.ToString(dr("C_LINKURL2"))
        txtLINKURL3.Text = Convert.ToString(dr("C_LINKURL3"))
        txtLINKURL4.Text = Convert.ToString(dr("C_LINKURL4"))
        txtLINKURL5.Text = Convert.ToString(dr("C_LINKURL5"))
        txtCSORT1.Text = Convert.ToString(dr("C_SORT1"))
        If flag_ROC Then
            txtEDATE1.Text = TIMS.Cdate17(Convert.ToString(dr("CEDATED")))  'edit，by:20181019
        Else
            txtEDATE1.Text = Convert.ToString(dr("CEDATED"))
        End If
        Common.SetListItem(ddlC_EDATE_hh1, Convert.ToString(dr("CEDATEHH")))
        Common.SetListItem(ddlC_EDATE_mm1, Convert.ToString(dr("CEDATEMM")))
        '==========
        If Convert.ToString(dr("CEDATEHH")) = "" Then ddlC_EDATE_hh1.SelectedIndex = -1
        If Convert.ToString(dr("CEDATEMM")) = "" Then ddlC_EDATE_mm1.SelectedIndex = -1
    End Sub

    '(2)讀取段落內容區塊
    '2019-01-16 ADD SEC_PICTURE_ALT 圖檔提示文字
    Private Sub Load_TB_CONTENT_SECTION(ByVal iSEQNO As Integer)
        If Convert.ToString(iSEQNO) = "" Then Exit Sub
        If iSEQNO = 0 Then Exit Sub

        Dim sql2 As String = ""
        sql2 &= " SELECT SEC_NO id" & vbCrLf
        sql2 &= " ,'區塊'+convert(varchar,SEC_NO) item" & vbCrLf
        sql2 &= " ,SEQNO" & vbCrLf
        sql2 &= " , CONTENTID" & vbCrLf
        sql2 &= " , SEC_NO" & vbCrLf
        sql2 &= " , SEC_CONTENT" & vbCrLf
        sql2 &= " , SEC_PICTURE" & vbCrLf
        sql2 &= " ,ALIGN_TYPE" & vbCrLf
        sql2 &= " , MODIFYACCT" & vbCrLf
        sql2 &= " , MODIFYDATE" & vbCrLf
        sql2 &= " ,SEC_PICTURE_ALT" & vbCrLf
        sql2 &= " FROM TB_CONTENT_SECTION" & vbCrLf
        sql2 &= " WHERE CONTENTID=@CONTENTID" & vbCrLf
        sql2 &= " ORDER BY CONTENTID ASC, SEC_NO ASC" & vbCrLf

        Dim parms2 As Hashtable = New Hashtable()
        If Convert.ToString(iSEQNO) <> "" Then parms2.Add("CONTENTID", iSEQNO)
        Dim dt2 As DataTable
        dt2 = DbAccess.GetDataTable(sql2, objconn, parms2)
        gv1.Visible = False '查無資料，不顯示
        If dt2.Rows.Count = 0 Then Exit Sub

        gv1.Visible = True
        gv1.DataSource = dt2
        gv1.DataBind()

        'For i As Integer=0 To dt2.Rows.Count - 1
        '    Dim dr As DataRow=dt2.Rows(i)
        '    Dim secNo As Integer=Convert.ToInt32(dr("SEC_NO"))
        '    Dim secNo_sys As Integer=secNo - 1
        '    CType(gv1.Rows(secNo_sys).FindControl("txtData"), TextBox).Text=Convert.ToString(dr("SEC_CONTENT"))
        '    CType(gv1.Rows(secNo_sys).FindControl("lblSEQ"), Label).Text=Convert.ToString(dr("SEQNO"))
        '    Dim picName As String=Convert.ToString(dr("SEC_PICTURE"))
        '    If picName <> "" Then
        '        CType(gv1.Rows(secNo_sys).FindControl("imgFUrl"), WebControls.Image).ImageUrl=ImageUrl_Path1 + picName
        '        CType(gv1.Rows(secNo_sys).FindControl("lblFName"), Label).Text=picName
        '        CType(gv1.Rows(secNo_sys).FindControl("divPic"), HtmlControls.HtmlGenericControl).Visible=True
        '    Else
        '        CType(gv1.Rows(secNo_sys).FindControl("divPic"), HtmlControls.HtmlGenericControl).Visible=False
        '    End If
        '    Dim rblPosition As RadioButtonList=CType(gv1.Rows(secNo_sys).FindControl("rblPosition"), RadioButtonList)
        '    Common.SetListItem(rblPosition, Convert.ToString(dr("ALIGN_TYPE")))
        '    'CType(gv1.Rows(secNo_sys).FindControl("rblPosition"), RadioButtonList).SelectedValue=Convert.ToString(dr("ALIGN_TYPE"))
        'Next

    End Sub

    '(3)讀取附件內容區塊
    Sub Load_TB_FILE()

        divFile.Visible = False
        If hid_f_grp.Value = "" Then Exit Sub

        Dim sql3 As String = ""
        sql3 &= " Select FILEID" & vbCrLf
        sql3 &= " , F_GROUPID, F_FUNID" & vbCrLf
        sql3 &= " , FILE_ORINAME, FILE_PHYNAME" & vbCrLf
        sql3 &= " , FILE_TYPE, FILE_SIZE, FILE_NO" & vbCrLf
        sql3 &= " , DLCOUNT, ISUSED" & vbCrLf
        sql3 &= " , MODIFYACCT, MODIFYDATE" & vbCrLf
        sql3 &= " FROM TB_FILE" & vbCrLf
        sql3 &= " WHERE ISUSED='Y'" & vbCrLf
        sql3 &= " AND F_GROUPID=@F_GROUPID" & vbCrLf
        sql3 &= " ORDER BY F_GROUPID ASC, FILE_NO ASC" & vbCrLf
        Dim parms3 As New Hashtable()
        parms3.Add("F_GROUPID", hid_f_grp.Value)
        Dim dt3 As DataTable
        dt3 = DbAccess.GetDataTable(sql3, objconn, parms3)

        gv2.DataSource = dt3
        gv2.DataBind()

        divFile.Visible = False
        If dt3.Rows.Count = 0 Then Exit Sub
        divFile.Visible = True

        'For i As Integer=0 To dt3.Rows.Count - 1
        '    Dim dr As DataRow=dt3.Rows(i)
        '    Dim F_Name1 As String=Convert.ToString(dr("FILE_ORINAME"))
        '    Dim F_Name2 As String=Convert.ToString(dr("FILE_PHYNAME"))
        '    Dim F_Type As String=Convert.ToString(dr("FILE_TYPE"))
        '    Dim F_Size As String=Convert.ToString(dr("FILE_SIZE"))

        '    CType(gv2.Rows(i).FindControl("lnkFName"), LinkButton).Text=F_Name1
        '    CType(gv2.Rows(i).FindControl("lnkFName"), LinkButton).ToolTip="[ 附檔名：" + F_Type.Trim.ToLower.Substring(1) + "  /  檔案大小：" + F_Size + " ]"
        '    CType(gv2.Rows(i).FindControl("lblFName"), Label).Text=F_Name2.Trim
        '    CType(gv2.Rows(i).FindControl("lblFExt"), Label).Text=F_Type.Trim.ToUpper
        '    CType(gv2.Rows(i).FindControl("lblFFileid"), Label).Text=Convert.ToString(dr("FILEID"))
        'Next

    End Sub

    '上傳段落圖片
    Protected Sub bt_upPic_Click(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        Dim gvr As GridViewRow = CType(btn.NamingContainer, GridViewRow)

        Dim fu1 As FileUpload = CType(gvr.FindControl("fu1"), FileUpload)
        Dim imgFUrl As WebControls.Image = CType(gvr.FindControl("imgFUrl"), WebControls.Image)
        Dim lblFName As Label = CType(gvr.FindControl("lblFName"), Label)

#Region "上傳段落圖片時,所需的變數"

        'Dim V_TYPE As String=TIMS.ClearSQM(ddlType.SelectedValue)
        Dim v_ddlType As String = TIMS.GetListValue(ddlType)

        Dim blHasNewP1 As Boolean = fu1.HasFile
        Dim blPassP1 As Boolean = False   '驗證[附件1]是否通過OpenOffice條件限制
        Dim ErrMsg As String = ""
        Dim Pic1Ext As String = ""
        Dim myType As String = "N" + v_ddlType + "_P_"
        Dim PoName As String = ""
        Dim PnName As String = ""
        Dim PoSize As Integer = 0
        Dim PnSize As String = ""
        Dim PnTime As String = DateTime.Now.ToString("yyyyMMddHHmmss")
        Dim pnPath As String = ""

#End Region

        If blHasNewP1 Then
            If fu1.FileBytes.Length > 10485760 Then ErrMsg &= "檔案大小超過10MB" & vbCrLf
            '==========
            Dim UpFile1 As HttpPostedFile = fu1.PostedFile
            Dim File1Split() As String = Split(UpFile1.FileName, "\")
            Dim File1Name As String = File1Split(File1Split.Length - 1)
            Pic1Ext = System.IO.Path.GetExtension(File1Name).ToUpper
            If CheckPicType(Pic1Ext) Then blPassP1 = True Else blPassP1 = False
            If Not blPassP1 Then ErrMsg &= "檔案類型必須是圖片檔(例如：jpg、jpeg、png、gif)" & vbCrLf
            If ErrMsg <> "" Then Common.MessageBox(Me, ErrMsg)
            If ErrMsg <> "" Then Exit Sub
            '==========
            PoSize = fu1.FileBytes.Length
            PnSize = GetFileSize(PoSize)
            PoName = File1Name
            PnName = myType + PnTime + Pic1Ext
            '==========
            If Not System.IO.Directory.Exists(Server.MapPath(mySavePath)) Then
                System.IO.Directory.CreateDirectory(Server.MapPath(mySavePath))
            End If
            pnPath = Server.MapPath(mySavePath + PnName)
            fu1.SaveAs(pnPath)
            '==========
            imgFUrl.ImageUrl = ImageUrl_Path1 + PnName
            lblFName.Text = PnName
            CType(gvr.FindControl("divPic"), HtmlControls.HtmlGenericControl).Visible = True
        End If
    End Sub

    '刪除段落圖片
    Protected Sub bt_delPic_Click(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        Dim gvr As GridViewRow = CType(btn.NamingContainer, GridViewRow)
        Dim imgFUrl As WebControls.Image = CType(gvr.FindControl("imgFUrl"), WebControls.Image)
        Dim lblFName As Label = CType(gvr.FindControl("lblFName"), Label)

        Dim pSavePath As String = Server.MapPath(mySavePath + lblFName.Text)
        If System.IO.File.Exists(pSavePath) Then System.IO.File.Delete(pSavePath)
        imgFUrl.ImageUrl = ""
        lblFName.Text = ""
        CType(gvr.FindControl("divPic"), HtmlControls.HtmlGenericControl).Visible = False
    End Sub

    '上傳附件檔案
    Protected Sub bt_upfile_Click(sender As Object, e As EventArgs) Handles bt_upfile.Click
#Region "[附件檔案]相關參數設定"

        Dim blHasNewF1 As Boolean = fu2.HasFile
        Dim blPassF1 As Boolean = False   '驗證[附件檔案]是否通過副檔名條件限制
        Dim ErrMsg As String = ""

        Dim UpFile As HttpPostedFile = fu2.PostedFile
        Dim s_FoName As String = ""
        Dim FExt As String = ""
        Dim s_FnName As String = ""
        Dim FPath As String = ""
        Dim FoSize As Integer = 0
        Dim FnSize As String = ""
        Dim FnTime As String = DateTime.Now.ToString("yyyyMMddHHmmss")

        If blHasNewF1 Then
            If fu2.FileBytes.Length > 10485760 Then ErrMsg &= "檔案大小超過10MB" & vbCrLf
            '==========
            FoSize = fu2.FileBytes.Length
            FnSize = GetFileSize(FoSize)
            Dim FileSplit() As String = Split(UpFile.FileName, "\")
            s_FoName = FileSplit(FileSplit.Length - 1)
            FExt = System.IO.Path.GetExtension(s_FoName).ToUpper
            If CheckFileType(FExt) Then blPassF1 = True Else blPassF1 = False
            If Not blPassF1 Then ErrMsg &= "目前可接受的檔案類型為doc、docx、ppt、pptx、xls、xlsx、odt、odp、ods、pdf" & vbCrLf
            If ErrMsg <> "" Then Common.MessageBox(Me, ErrMsg)
            If ErrMsg <> "" Then Exit Sub

            '==========
            'Dim V_TYPE As String=TIMS.ClearSQM(ddlType.SelectedValue)
            Dim v_ddlType As String = TIMS.GetListValue(ddlType)
            s_FnName = "N" + v_ddlType + "_F_" + FnTime + FExt
            FPath = Server.MapPath(mySavePath + s_FnName)
            fu2.SaveAs(FPath)
        End If

#End Region

        If blPassF1 And ErrMsg = "" Then
            Call TIMS.OpenDbConn(objconn)

            '==========
            Dim addSql As String = ""
            addSql &= " INSERT INTO TB_FILE (FILEID, F_GROUPID, F_FUNID, FILE_ORINAME, FILE_PHYNAME, FILE_TYPE, FILE_SIZE, FILE_NO, DLCOUNT, MODIFYACCT, MODIFYDATE)" & vbCrLf
            addSql &= " VALUES (@FILEID, @F_GROUPID, @F_FUNID, @FILE_ORINAME, @FILE_PHYNAME, @FILE_TYPE, @FILE_SIZE, @FILE_NO, 0, @MODIFYACCT, GETDATE())" & vbCrLf

            Dim tFileId As Integer = DbAccess.GetNewId(objconn, "TB_FILE_FILEID_SEQ,TB_FILE,FILEID")
            '==========
            If hid_f_grp.Value = "" Then
                hid_f_grp.Value = Get_F_GROUPID()
                'Dim selSql As String="" & vbCrLf
                'selSql &= " SELECT MAX(ISNULL(F_GROUPID, 0)) AS myNowId FROM TB_FILE "
                'Dim tDt As DataTable=DbAccess.GetDataTable(selSql, objconn)
                'hid_f_grp.Value="1"
                'If tDt.Rows.Count > 0 Then
                '    Dim tDr As DataRow=tDt.Rows(0)
                '    hid_f_grp.Value=Convert.ToString(Convert.ToInt32(tDr("myNowId")) + 1).Trim
                'End If
            End If

            Dim v_ddlType As String = TIMS.GetListValue(ddlType)
            Dim addParms As New Hashtable From {
                {"FILEID", tFileId},
                {"F_GROUPID", Convert.ToInt32(hid_f_grp.Value)},
                {"F_FUNID", v_ddlType},
                {"FILE_ORINAME", s_FoName.Trim},
                {"FILE_PHYNAME", s_FnName.Trim},
                {"FILE_TYPE", FExt.Trim},
                {"FILE_SIZE", FnSize.Trim},
                {"FILE_NO", (gv2.Rows.Count + 1).ToString()},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            Dim iRst As Integer = DbAccess.ExecuteNonQuery(addSql, objconn, addParms)

            '重新取得DB的附件資訊
            'show hid_f_grp.Value
            Load_TB_FILE()
        End If
    End Sub

    '將附件檔案下載到客戶端
    Protected Sub lnkFName_Click(sender As Object, e As EventArgs)
        Dim lnkBtn As LinkButton = CType(sender, LinkButton)
        Dim gvr As GridViewRow = CType(lnkBtn.NamingContainer, GridViewRow)
        Dim lblFName As Label = CType(gvr.FindControl("lblFName"), Label)
        Dim filePath1 As String = mySavePath + lblFName.Text
        Dim filePath2 As String = Server.MapPath(filePath1)

        If Not System.IO.File.Exists(filePath2) Then
            Common.MessageBox(Page, "附件檔案已不存在，請重新上傳，謝謝。")
            Exit Sub
        End If

        Dim fileByte As Byte() = System.IO.File.ReadAllBytes(filePath2)
        Dim myNewFileName As String = lnkBtn.Text
        Response.Clear()
        Response.AddHeader("Content-Disposition", "attachment;filename=" + myNewFileName)
        Response.BinaryWrite(fileByte)
        Response.End()
    End Sub

    '刪除附件檔案
    Protected Sub bt_delFile_Click(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        Dim gvr As GridViewRow = CType(btn.NamingContainer, GridViewRow)
        Dim lblFFileid As Label = CType(gvr.FindControl("lblFFileid"), Label)

        Call TIMS.OpenDbConn(objconn)

        '刪除DB的附件資訊
        Dim delSql As String = " DELETE FROM TB_FILE WHERE FILEID=@FILEID" & vbCrLf
        Dim delParms As New Hashtable From {
            {"FILEID", lblFFileid.Text.Trim}
        }
        Dim iRst As Integer = DbAccess.ExecuteNonQuery(delSql, objconn, delParms)

        '刪除附件檔案
        Dim lnkFName As LinkButton = CType(gvr.FindControl("lnkFName"), LinkButton)
        Dim lblFName As Label = CType(gvr.FindControl("lblFName"), Label)
        Dim bt_delFile As Button = CType(gvr.FindControl("bt_delFile"), Button)
        Dim fSavePath As String = Server.MapPath(mySavePath + lblFName.Text)
        If System.IO.File.Exists(fSavePath) Then System.IO.File.Delete(fSavePath)
        lnkFName.Text = ""
        lnkFName.Enabled = False
        lblFName.Text = ""
        bt_delFile.Visible = False

        '重新取得DB的附件資訊
        'show hid_f_grp.Value
        Load_TB_FILE()

        '重新調整剩餘附件的序號

        'Dim i As Integer
        For i As Integer = 0 To gv2.Rows.Count - 1
            Dim tFFileid As Label = CType(gv2.Rows(i).FindControl("lblFFileid"), Label)
            Dim upSql As String = " UPDATE TB_FILE SET FILE_NO=@FILE_NO WHERE FILEID= @FILEID" & vbCrLf
            Dim upParms As New Hashtable From {
                {"FILE_NO", (i + 1).ToString()},
                {"FILEID", tFFileid.Text.Trim}
            }
            Dim tRst As Integer = DbAccess.ExecuteNonQuery(upSql, objconn, upParms)
        Next
    End Sub

    '進行「儲存」作業
    Protected Sub bt_save_Click(sender As Object, e As EventArgs) Handles bt_save.Click
        Dim Errmsg As String = ""
        Call CheckData(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        Call SaveData1()
    End Sub

    '送出前的檢核
    Function CheckData(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True
        Errmsg = ""

        If ddlType.SelectedIndex = -1 Then Errmsg &= "類別 不可為空" & vbCrLf

        txtSDATE1.Text = TIMS.ClearSQM(txtSDATE1.Text)  'edit，by:20181019
        txtEDATE1.Text = TIMS.ClearSQM(txtEDATE1.Text)  'edit，by:20181019
        Dim mySDATE1 As String = If(flag_ROC, TIMS.Cdate18(txtSDATE1.Text), txtSDATE1.Text).Replace("/", "-")  'edit，by:20181019
        Dim myEDATE1 As String = If(flag_ROC, TIMS.Cdate18(txtEDATE1.Text), txtEDATE1.Text).Replace("/", "-")  'edit，by:20181019
        txtTitle.Text = TIMS.ClearSQM(txtTitle.Text)
        txtLINKURL1.Text = TIMS.ClearSQM(txtLINKURL1.Text)
        txtLINKURL2.Text = TIMS.ClearSQM(txtLINKURL2.Text)
        txtLINKURL3.Text = TIMS.ClearSQM(txtLINKURL3.Text)
        txtLINKURL4.Text = TIMS.ClearSQM(txtLINKURL4.Text)
        txtLINKURL5.Text = TIMS.ClearSQM(txtLINKURL5.Text)

        Dim oC_SDATE As Object = mySDATE1 + " " + ddlC_SDATE_hh1.SelectedValue + ":" + ddlC_SDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019
        Dim oC_EDATE As Object = myEDATE1 + " " + ddlC_EDATE_hh1.SelectedValue + ":" + ddlC_EDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019

        If txtSDATE1.Text = "" OrElse TIMS.CStr1(oC_SDATE) = "" Then Errmsg &= "上架日期 不可為空" & vbCrLf
        If txtTitle.Text = "" Then Errmsg &= "標題 不可為空" & vbCrLf
        txtCSORT1.Text = TIMS.ClearSQM(txtCSORT1.Text)
        txtCSORT1.Text = TIMS.ChangeIDNO(txtCSORT1.Text)
        If txtCSORT1.Text <> "" Then
            If Not TIMS.IsNumeric1(txtCSORT1.Text) Then Errmsg &= "序號 不可為非數字" & vbCrLf
        End If

        Dim v_ddlType As String = TIMS.GetListValue(ddlType)
        Select Case v_ddlType
            Case "011"
                txtC_URL.Text = TIMS.ClearSQM(txtC_URL.Text)
                txtC_CONTENT1.Text = TIMS.ClearSQM(txtC_CONTENT1.Text)
                If txtC_URL.Text = "" Then Errmsg &= "宣導影片網址 不可為空" & vbCrLf
                If txtC_CONTENT1.Text = "" Then Errmsg &= "宣導連結 不可為空" & vbCrLf
        End Select

        '檢核圖檔提示文字必填與否
        Dim lblFName As Label = Nothing
        Dim txtPicAlt As TextBox = Nothing

        If gv1 IsNot Nothing Then
            For i As Integer = 0 To gv1.Rows.Count - 1
                Dim gvr As GridViewRow = gv1.Rows(i)
                lblFName = gvr.FindControl("lblFName")
                txtPicAlt = gvr.FindControl("txtPicAlt")

                If lblFName.Text <> "" AndAlso txtPicAlt.Text.Trim = "" Then
                    Errmsg &= "區塊 " & (i + 1).ToString() & " 圖檔提示文字 不可為空" & vbCrLf
                End If
            Next
        End If

        If txtEDATE1.Text = "" OrElse TIMS.CStr1(oC_EDATE) = "" Then Errmsg &= "停用日期 不可為空" & vbCrLf
        If Errmsg <> "" Then Return False

        If DateDiff(DateInterval.Minute, CDate(oC_SDATE), CDate(oC_EDATE)) = 0 Then
            Errmsg &= "上架日期與停用日期 不可相等!!" & vbCrLf
        End If
        If DateDiff(DateInterval.Minute, CDate(oC_SDATE), CDate(oC_EDATE)) < 0 Then
            Errmsg &= "上架日期與停用日期 順序異常!!" & vbCrLf
        End If
        If Errmsg <> "" Then Return False

        '附件檢核 ============================ Start
        If gv2 IsNot Nothing Then
            If gv2.Rows.Count > 0 Then
                Dim blOk As Boolean = False
                For i As Integer = 0 To gv2.Rows.Count - 1
                    Dim F_Type As String = Convert.ToString(CType(gv2.Rows(i).FindControl("lblFExt"), Label).Text.Trim)
                    blOk = CheckOpenFileType(F_Type.ToUpper)
                    If blOk = True Then Exit For
                Next
                If Not blOk Then Errmsg &= "至少上傳1個含有ODF類型(例如：odp、ods、odt、pdf)的檔案!"
            End If
        End If
        '=================================== End

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    '儲存(part-1)
    Sub SaveData1()
        Dim flagSaveOK1 As Boolean = False
        Try
            flagSaveOK1 = SaveData1_2()
        Catch ex As Exception
            flagSaveOK1 = False
            Common.MessageBox(Me, ex.Message)
            Exit Sub
        End Try

        If flagSaveOK1 Then
            '儲存成功
            Dim url1 As String = "RWB_01_001.aspx?id1=" & TIMS.Get_MRqID(Me)
            'Common.MessageBox(Me, "儲存成功!", url1)
            Common.MessageBox(Me, "儲存成功!")
            TIMS.Utl_Redirect(Me, objconn, url1)
        End If
    End Sub

    ''' <summary>
    ''' [基本內容區塊]儲存
    ''' </summary>
    Sub Save_CONTENT(ByRef saveMode As String)
        '[基本內容區塊]儲存
        txtSDATE1.Text = TIMS.ClearSQM(txtSDATE1.Text)
        txtEDATE1.Text = TIMS.ClearSQM(txtEDATE1.Text)
        Dim mySDATE1 As String = If(flag_ROC, TIMS.Cdate18(txtSDATE1.Text), txtSDATE1.Text).Replace("/", "-")  'edit，by:20181019
        Dim myEDATE1 As String = If(flag_ROC, TIMS.Cdate18(txtEDATE1.Text), txtEDATE1.Text).Replace("/", "-")  'edit，by:20181019
        txtTitle.Text = TIMS.ClearSQM(txtTitle.Text)
        txtLINKURL1.Text = TIMS.ClearSQM(txtLINKURL1.Text)
        txtLINKURL2.Text = TIMS.ClearSQM(txtLINKURL2.Text)
        txtLINKURL3.Text = TIMS.ClearSQM(txtLINKURL3.Text)
        txtLINKURL4.Text = TIMS.ClearSQM(txtLINKURL4.Text)
        txtLINKURL5.Text = TIMS.ClearSQM(txtLINKURL5.Text)
        Dim oC_SDATE As Object = mySDATE1 + " " + ddlC_SDATE_hh1.SelectedValue + ":" + ddlC_SDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019
        Dim oC_EDATE As Object = myEDATE1 + " " + ddlC_EDATE_hh1.SelectedValue + ":" + ddlC_EDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019
        'Dim vSUB_FUNID As String=TIMS.ClearSQM(ddlType.SelectedValue)
        Dim vSUB_FUNID As String = TIMS.GetListValue(ddlType)

        Dim iSql As String = ""
        iSql &= " INSERT INTO TB_CONTENT(SEQNO, FUNID, C_SDATE, C_EDATE, C_TITLE,C_LINKURL1,C_LINKURL2, C_CDATE, C_CACCT, C_UDATE, C_UACCT, C_STATUS, SUB_FUNID, F_GROUPID,C_SORT1)" & vbCrLf
        iSql &= " VALUES (@SEQNO, @FUNID, @C_SDATE, @C_EDATE, @C_TITLE,@C_LINKURL1,@C_LINKURL2, GETDATE(), @C_CACCT, GETDATE(), @C_UACCT, 'A', @SUB_FUNID, @F_GROUPID,@C_SORT1)" & vbCrLf
        'Dim iCmd As New SqlCommand(iSql, objconn)

        Dim uSql As String = ""
        uSql &= " UPDATE TB_CONTENT" & vbCrLf
        uSql &= " SET C_SDATE=@C_SDATE" & vbCrLf
        uSql &= "  ,C_EDATE=@C_EDATE" & vbCrLf
        uSql &= "  ,C_TITLE=@C_TITLE" & vbCrLf
        uSql &= "  ,C_LINKURL1=@C_LINKURL1" & vbCrLf
        uSql &= "  ,C_LINKURL2=@C_LINKURL2" & vbCrLf
        uSql &= "  ,C_UDATE=GETDATE()" & vbCrLf
        uSql &= "  ,C_UACCT=@C_UACCT" & vbCrLf
        uSql &= "  ,C_STATUS='M'" & vbCrLf
        uSql &= "  ,SUB_FUNID=@SUB_FUNID" & vbCrLf
        uSql &= "  ,F_GROUPID=@F_GROUPID" & vbCrLf
        uSql &= "  ,C_SORT1=@C_SORT1" & vbCrLf
        uSql &= " WHERE SEQNO=@SEQNO" & vbCrLf
        'Dim uCmd As New SqlCommand(uSql, objconn)

        If hid_f_grp.Value = "" Then hid_f_grp.Value = Get_F_GROUPID()

        Call TIMS.OpenDbConn(objconn)
        Dim iRst As Integer = 0
        If hid_V_SEQNO.Value = "" Then
            '新增
            Dim iSEQNO As Integer = DbAccess.GetNewId(objconn, "TB_CONTENT_SEQNO_SEQ,TB_CONTENT,SEQNO")
            Dim parms As New Hashtable()
            parms.Add("SEQNO", iSEQNO)
            parms.Add("FUNID", "001")
            parms.Add("C_SDATE", oC_SDATE)
            parms.Add("C_EDATE", oC_EDATE)
            parms.Add("C_TITLE", txtTitle.Text)
            parms.Add("C_LINKURL1", If(txtLINKURL1.Text <> "", txtLINKURL1.Text, Convert.DBNull))
            parms.Add("C_LINKURL2", If(txtLINKURL2.Text <> "", txtLINKURL2.Text, Convert.DBNull))
            parms.Add("C_LINKURL3", If(txtLINKURL3.Text <> "", txtLINKURL3.Text, Convert.DBNull))
            parms.Add("C_LINKURL4", If(txtLINKURL4.Text <> "", txtLINKURL4.Text, Convert.DBNull))
            parms.Add("C_LINKURL5", If(txtLINKURL5.Text <> "", txtLINKURL5.Text, Convert.DBNull))
            parms.Add("C_CACCT", sm.UserInfo.UserID)
            parms.Add("C_UACCT", sm.UserInfo.UserID)
            parms.Add("SUB_FUNID", vSUB_FUNID)
            parms.Add("F_GROUPID", hid_f_grp.Value)
            parms.Add("C_SORT1", If(txtCSORT1.Text <> "", Val(txtCSORT1.Text), iSEQNO))
            iRst += DbAccess.ExecuteNonQuery(iSql, objconn, parms)
            hid_V_SEQNO.Value = iSEQNO
            saveMode = "I"
        Else
            '修改
            Dim parms As Hashtable = New Hashtable()
            parms.Add("SEQNO", hid_V_SEQNO.Value)
            parms.Add("C_SDATE", oC_SDATE)
            parms.Add("C_EDATE", oC_EDATE)
            parms.Add("C_TITLE", txtTitle.Text)
            parms.Add("C_LINKURL1", If(txtLINKURL1.Text <> "", txtLINKURL1.Text, Convert.DBNull))
            parms.Add("C_LINKURL2", If(txtLINKURL2.Text <> "", txtLINKURL2.Text, Convert.DBNull))
            parms.Add("C_LINKURL3", If(txtLINKURL3.Text <> "", txtLINKURL3.Text, Convert.DBNull))
            parms.Add("C_LINKURL4", If(txtLINKURL4.Text <> "", txtLINKURL4.Text, Convert.DBNull))
            parms.Add("C_LINKURL5", If(txtLINKURL5.Text <> "", txtLINKURL5.Text, Convert.DBNull))
            parms.Add("C_UACCT", sm.UserInfo.UserID)
            parms.Add("SUB_FUNID", vSUB_FUNID)
            parms.Add("F_GROUPID", hid_f_grp.Value)
            parms.Add("C_SORT1", If(txtCSORT1.Text <> "", Val(txtCSORT1.Text), Val(hid_V_SEQNO.Value)))
            iRst += DbAccess.ExecuteNonQuery(uSql, objconn, parms)
            saveMode = "M"
        End If
    End Sub

    ''' <summary>
    ''' [段落內容區塊]儲存
    ''' </summary>
    Sub Save_SECTION(ByRef saveMode As String)

        Dim iSql2 As String = ""
        iSql2 &= " INSERT INTO TB_CONTENT_SECTION (SEQNO, CONTENTID, SEC_NO, SEC_CONTENT, SEC_PICTURE, SEC_PICTURE_ALT, ALIGN_TYPE, MODIFYACCT, MODIFYDATE)" & vbCrLf
        iSql2 &= " VALUES (@SEQNO, @CONTENTID, @SEC_NO, @SEC_CONTENT, @SEC_PICTURE, @SEC_PICTURE_ALT,@ALIGN_TYPE, @MODIFYACCT, GETDATE())" & vbCrLf
        'Dim iCmd2 As New SqlCommand(iSql2, objconn)

        Dim uSql2 As String = ""
        uSql2 &= " UPDATE TB_CONTENT_SECTION" & vbCrLf
        uSql2 &= " SET CONTENTID=@CONTENTID" & vbCrLf
        uSql2 &= " ,SEC_NO=@SEC_NO" & vbCrLf
        uSql2 &= " ,SEC_CONTENT=@SEC_CONTENT" & vbCrLf
        uSql2 &= " ,SEC_PICTURE=@SEC_PICTURE" & vbCrLf
        uSql2 &= " ,SEC_PICTURE_ALT=@SEC_PICTURE_ALT" & vbCrLf
        uSql2 &= " ,ALIGN_TYPE=@ALIGN_TYPE" & vbCrLf
        uSql2 &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        uSql2 &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        uSql2 &= " WHERE SEQNO=@SEQNO" & vbCrLf
        'Dim uCmd2 As New SqlCommand(uSql2, objconn)

        Call TIMS.OpenDbConn(objconn)

        '[段落內容區塊]儲存
        'Dim i As Integer
        If gv1 Is Nothing Then Exit Sub
        If gv1.Rows.Count = 0 Then Exit Sub

        For i As Integer = 0 To i_contentCount - 1
            Dim txtPicAlt As TextBox = gv1.Rows(i).FindControl("txtPicAlt")
            Dim lblFName As Label = gv1.Rows(i).FindControl("lblFName")
            Dim lblSEQ As Label = gv1.Rows(i).FindControl("lblSEQ")
            Dim lblSecNo As Label = gv1.Rows(i).FindControl("lblSecNo")
            Dim txtData As TextBox = gv1.Rows(i).FindControl("txtData")
            Dim rblPosition As RadioButtonList = gv1.Rows(i).FindControl("rblPosition")
            lblSEQ.Text = TIMS.ClearSQM(lblSEQ.Text)
            Dim v_rblPosition As String = TIMS.GetListValue(rblPosition)

            Dim iRst2 As Integer = 0
            If saveMode.Equals("I") Or (saveMode.Equals("M") And lblSEQ.Text = "") Then  'edit，by:20181019
                '新增
                Dim iSEQNO2 As Integer = DbAccess.GetNewId(objconn, "TB_CONTENT_SECTION_SEQNO,TB_CONTENT_SECTION,SEQNO")
                'With iCmd2
                Dim parms As Hashtable = New Hashtable()
                parms.Add("SEQNO", iSEQNO2)
                parms.Add("CONTENTID", hid_V_SEQNO.Value)
                parms.Add("SEC_NO", lblSecNo.Text)
                parms.Add("SEC_CONTENT", txtData.Text)
                '==========
                parms.Add("SEC_PICTURE", If(lblFName.Text <> "", lblFName.Text, Convert.DBNull))
                '圖檔文字說明
                parms.Add("SEC_PICTURE_ALT", If(txtPicAlt.Text <> "", txtPicAlt.Text, Convert.DBNull))
                '==========
                parms.Add("ALIGN_TYPE", If(v_rblPosition <> "", v_rblPosition, Convert.DBNull))
                parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                iRst2 += DbAccess.ExecuteNonQuery(iSql2, objconn, parms)
                'End With
                lblSEQ.Text = iSEQNO2
            Else
                '修改
                'With uCmd2
                Dim parms As Hashtable = New Hashtable()
                parms.Add("SEQNO", lblSEQ.Text)  'edit，by:20181019
                parms.Add("CONTENTID", hid_V_SEQNO.Value)
                parms.Add("SEC_NO", lblSecNo.Text)
                parms.Add("SEC_CONTENT", txtData.Text)
                '==========
                parms.Add("SEC_PICTURE", If(lblFName.Text <> "", lblFName.Text, Convert.DBNull))
                '圖檔文字說明
                parms.Add("SEC_PICTURE_ALT", If(txtPicAlt.Text <> "", txtPicAlt.Text, Convert.DBNull))
                '==========
                parms.Add("ALIGN_TYPE", If(v_rblPosition <> "", v_rblPosition, Convert.DBNull))
                parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                iRst2 += DbAccess.ExecuteNonQuery(uSql2, objconn, parms)
                'End With
            End If
        Next

    End Sub

    ''' <summary>
    '''  [基本內容區塊]儲存 依 s_FUNID
    ''' </summary>
    ''' <param name="s_FUNID"></param>
    Sub Save_CONTENT_FUN(ByVal s_FUNID As String)
        'Const s_FUNID As String="011"
        '[基本內容區塊]儲存
        txtSDATE1.Text = TIMS.ClearSQM(txtSDATE1.Text)
        txtEDATE1.Text = TIMS.ClearSQM(txtEDATE1.Text)
        Dim mySDATE1 As String = If(flag_ROC, TIMS.Cdate18(txtSDATE1.Text), txtSDATE1.Text).Replace("/", "-")  'edit，by:20181019
        Dim myEDATE1 As String = If(flag_ROC, TIMS.Cdate18(txtEDATE1.Text), txtEDATE1.Text).Replace("/", "-")  'edit，by:20181019
        txtTitle.Text = TIMS.ClearSQM(txtTitle.Text)
        txtLINKURL1.Text = TIMS.ClearSQM(txtLINKURL1.Text)
        txtLINKURL2.Text = TIMS.ClearSQM(txtLINKURL2.Text)
        txtLINKURL3.Text = TIMS.ClearSQM(txtLINKURL3.Text)
        txtLINKURL4.Text = TIMS.ClearSQM(txtLINKURL4.Text)
        txtLINKURL5.Text = TIMS.ClearSQM(txtLINKURL5.Text)
        Dim oC_SDATE As Object = mySDATE1 + " " + ddlC_SDATE_hh1.SelectedValue + ":" + ddlC_SDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019
        Dim oC_EDATE As Object = myEDATE1 + " " + ddlC_EDATE_hh1.SelectedValue + ":" + ddlC_EDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019
        'Dim vSUB_FUNID As String=TIMS.ClearSQM(ddlType.SelectedValue)

        Dim iSql As String = ""
        iSql &= " INSERT INTO TB_CONTENT (SEQNO,FUNID,C_SDATE,C_EDATE,C_TITLE,C_URL,C_CONTENT1 ,C_CDATE,C_CACCT,C_UDATE,C_UACCT,C_STATUS, C_SORT1)" & vbCrLf
        iSql &= " VALUES (@SEQNO,@FUNID,@C_SDATE,@C_EDATE,@C_TITLE,@C_URL,@C_CONTENT1 ,GETDATE(),@C_CACCT,GETDATE(),@C_UACCT,'A', @C_SORT1)" & vbCrLf

        Dim uSql As String = ""
        uSql &= " UPDATE TB_CONTENT" & vbCrLf
        uSql &= " SET C_SDATE=@C_SDATE" & vbCrLf
        uSql &= " ,C_EDATE=@C_EDATE" & vbCrLf
        uSql &= " ,C_TITLE=@C_TITLE" & vbCrLf
        uSql &= " ,C_URL=@C_URL" & vbCrLf
        uSql &= " ,C_CONTENT1=@C_CONTENT1" & vbCrLf
        uSql &= " ,C_UDATE=GETDATE()" & vbCrLf
        uSql &= " ,C_UACCT=@C_UACCT" & vbCrLf
        uSql &= " ,C_STATUS='M'" & vbCrLf
        uSql &= " ,C_SORT1=@C_SORT1" & vbCrLf
        uSql &= " WHERE SEQNO=@SEQNO" & vbCrLf

        Call TIMS.OpenDbConn(objconn)

        Dim iRst As Integer = 0
        If hid_V_SEQNO.Value = "" Then
            '新增
            Dim iSEQNO As Integer = DbAccess.GetNewId(objconn, "TB_CONTENT_SEQNO_SEQ,TB_CONTENT,SEQNO")
            Dim parms As Hashtable = New Hashtable()
            parms.Add("SEQNO", iSEQNO)
            parms.Add("FUNID", s_FUNID)
            parms.Add("C_SDATE", oC_SDATE)
            parms.Add("C_EDATE", oC_EDATE)
            parms.Add("C_TITLE", txtTitle.Text)
            parms.Add("C_URL", txtC_URL.Text)
            parms.Add("C_CONTENT1", txtC_CONTENT1.Text)
            parms.Add("C_CACCT", sm.UserInfo.UserID)
            parms.Add("C_UACCT", sm.UserInfo.UserID)
            parms.Add("C_SORT1", If(txtCSORT1.Text <> "", Val(txtCSORT1.Text), iSEQNO))
            iRst += DbAccess.ExecuteNonQuery(iSql, objconn, parms)
            hid_V_SEQNO.Value = iSEQNO
            'saveMode="I"
        Else
            '修改
            Dim parms As Hashtable = New Hashtable()
            parms.Add("SEQNO", hid_V_SEQNO.Value)
            parms.Add("C_SDATE", oC_SDATE)
            parms.Add("C_EDATE", oC_EDATE)
            parms.Add("C_TITLE", txtTitle.Text)
            parms.Add("C_URL", txtC_URL.Text)
            parms.Add("C_CONTENT1", txtC_CONTENT1.Text)
            parms.Add("C_UACCT", sm.UserInfo.UserID)
            parms.Add("C_SORT1", If(txtCSORT1.Text <> "", Val(txtCSORT1.Text), Val(hid_V_SEQNO.Value)))
            iRst += DbAccess.ExecuteNonQuery(uSql, objconn, parms)
            'saveMode="M"
        End If
    End Sub

    ''' <summary>
    ''' 儲存(part-1-2)
    ''' </summary>
    ''' <returns></returns>
    Function SaveData1_2() As Boolean
        Dim rst As Boolean = False 'false:異常

        Dim v_ddlType As String = TIMS.GetListValue(ddlType)
        Select Case v_ddlType
            Case "1", "2", "3" '原本的儲存
                Dim saveMode As String = "" '傳遞之區塊儲存
                '[基本內容區塊]儲存
                Save_CONTENT(saveMode)
                '[段落內容區塊]儲存
                Save_SECTION(saveMode)

            Case "011"
                Save_CONTENT_FUN(v_ddlType)

            Case Else
                '新增的儲存
                Return rst
        End Select

        rst = True
        Return rst
    End Function

    ''' <summary>
    ''' 在[新增]模式裡,若沒有儲存時,則清除畫面上所有段落圖片-附件檔案
    ''' </summary>
    Sub DEL_FILE_ADDNEW()
        If hid_V_SEQNO.Value <> "" Then Exit Sub

        If gv1 Is Nothing Then Exit Sub
        If gv1.Rows.Count = 0 Then Exit Sub

        '刪除段落圖片
        'Dim i As Integer
        For i As Integer = 0 To i_contentCount - 1
            Dim imgFUrl As WebControls.Image = CType(gv1.Rows(i).FindControl("imgFUrl"), WebControls.Image)
            Dim lblFName As Label = CType(gv1.Rows(i).FindControl("lblFName"), Label)
            Dim pSavePath As String = Server.MapPath(mySavePath + lblFName.Text)
            If System.IO.File.Exists(pSavePath) Then System.IO.File.Delete(pSavePath)
            imgFUrl.ImageUrl = ""
            lblFName.Text = ""
            CType(gv1.Rows(i).FindControl("divPic"), HtmlControls.HtmlGenericControl).Visible = False
        Next

        '刪除DB的附件資訊
        'Dim j As Integer
        For j As Integer = 0 To gv2.Rows.Count - 1
            Dim lblFFileid As Label = CType(gv2.Rows(j).FindControl("lblFFileid"), Label)
            '==========
            Dim delSql As String = ""
            delSql &= " DELETE FROM TB_FILE WHERE FILEID=@FILEID" & vbCrLf
            Dim delParms As Hashtable = New Hashtable()
            delParms.Add("FILEID", lblFFileid.Text.Trim)
            Dim iRst As Integer = DbAccess.ExecuteNonQuery(delSql, objconn, delParms)
        Next

        '刪除附件檔案
        'Dim k As Integer
        For k As Integer = 0 To gv2.Rows.Count - 1
            Dim lnkFName As LinkButton = CType(gv2.Rows(k).FindControl("lnkFName"), LinkButton)
            Dim lblFName As Label = CType(gv2.Rows(k).FindControl("lblFName"), Label)
            Dim bt_delFile As Button = CType(gv2.Rows(k).FindControl("bt_delFile"), Button)
            Dim fSavePath As String = Server.MapPath(mySavePath + lblFName.Text)
            If System.IO.File.Exists(fSavePath) Then System.IO.File.Delete(fSavePath)
            lnkFName.Text = ""
            lnkFName.Enabled = False
            lblFName.Text = ""
            bt_delFile.Visible = False
        Next
    End Sub

    ''' <summary>
    ''' 在[修改]模式裡,若沒有儲存且[段落內容區塊]的圖片被刪除時,則更新DB資訊
    ''' </summary>
    Sub DEL_FILE_EDIT1()
        If hid_V_SEQNO.Value = "" Then Exit Sub

        If gv1 Is Nothing Then Exit Sub
        If gv1.Rows.Count = 0 Then Exit Sub

        'Dim i As Integer
        For i As Integer = 0 To i_contentCount - 1
            Dim lblFName As Label = CType(gv1.Rows(i).FindControl("lblFName"), Label)
            Dim lblSEQ As Label = CType(gv1.Rows(i).FindControl("lblSEQ"), Label)
            If lblSEQ Is Nothing Then Exit For
            lblSEQ.Text = TIMS.ClearSQM(lblSEQ.Text)
            If lblSEQ.Text = "" Then Exit For
            If Not TIMS.IsNumeric1(lblSEQ.Text) Then Exit For

            If lblFName.Text = "" Then
                Dim uSql As String = ""
                uSql &= " UPDATE TB_CONTENT_SECTION" & vbCrLf
                uSql &= " SET SEC_PICTURE=@SEC_PICTURE" & vbCrLf
                uSql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
                uSql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
                uSql &= " WHERE SEQNO=@SEQNO" & vbCrLf
                'Dim uCmd As New SqlCommand(uSql, objconn)
                Dim parms As Hashtable = New Hashtable()
                parms.Clear()
                parms.Add("SEQNO", lblSEQ.Text)  'edit，by:20181019
                parms.Add("SEC_PICTURE", Convert.DBNull)
                parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                Dim iRst As Integer = DbAccess.ExecuteNonQuery(uSql, objconn, parms)
            End If
        Next
    End Sub

    ''' <summary>
    ''' 進行「取消」作業
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub bt_cancle_Click(sender As Object, e As EventArgs) Handles bt_cancle.Click
        Call TIMS.OpenDbConn(objconn)
        '在[新增]模式裡,若沒有儲存時,則清除畫面上所有段落圖片&附件檔案
        If hid_V_SEQNO.Value = "" Then
            DEL_FILE_ADDNEW()
        End If
        '在[修改]模式裡,若沒有儲存且[段落內容區塊]的圖片被刪除時,則更新DB資訊
        If hid_V_SEQNO.Value <> "" Then
            DEL_FILE_EDIT1()
        End If

        Dim url1 As String = ""
        url1 = "RWB_01_001.aspx?id1=" & TIMS.Get_MRqID(Me)
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

#Region "檢查附件是否為圖片類型"

    '檢查附件是否為圖片類型
    Function CheckPicType(ByVal nowFileType As String) As Boolean
        Dim rst As Boolean = False

        '定義目前包含的圖片類型
        Dim acceptTypeList As New List(Of String)
        acceptTypeList.Add(".JPG")
        acceptTypeList.Add(".JPEG")
        acceptTypeList.Add(".PNG")
        acceptTypeList.Add(".GIF")

        'Dim i As Integer
        For i As Integer = 0 To acceptTypeList.Count - 1
            If nowFileType.Trim.ToUpper.Equals(acceptTypeList.Item(i)) Then
                rst = True
                Exit For
            End If
        Next

        Return rst
    End Function

#End Region

#Region "檢查附件是否包含Office的File類型"

    '檢查附件是否包含Office的File類型
    Function CheckFileType(ByVal nowFileType As String) As Boolean
        Dim rst As Boolean = False

        '定義目前Office包含的檔案類型
        Dim acceptTypeList As New List(Of String)
        acceptTypeList.Add(".DOC")
        acceptTypeList.Add(".DOCX")
        acceptTypeList.Add(".PPT")
        acceptTypeList.Add(".PPTX")
        acceptTypeList.Add(".XLS")
        acceptTypeList.Add(".XLSX")
        acceptTypeList.Add(".ODP")
        acceptTypeList.Add(".ODS")
        acceptTypeList.Add(".ODT")
        acceptTypeList.Add(".PDF")

        'Dim i As Integer
        For i As Integer = 0 To acceptTypeList.Count - 1
            If nowFileType.Trim.ToUpper.Equals(acceptTypeList.Item(i)) Then
                rst = True
                Exit For
            End If
        Next

        Return rst
    End Function

#End Region

#Region "檢查附件是否包含OpenOffile的File類型"

    '檢查附件是否包含OpenOffile的File類型
    Function CheckOpenFileType(ByVal nowFileType As String) As Boolean
        Dim rst As Boolean = False

        '定義目前OpenOffice包含的檔案類型
        Dim acceptTypeList As New List(Of String)
        acceptTypeList.Add(".ODP")
        acceptTypeList.Add(".ODS")
        acceptTypeList.Add(".ODT")
        acceptTypeList.Add(".PDF")

        'Dim i As Integer
        For i As Integer = 0 To acceptTypeList.Count - 1
            If nowFileType.Trim.ToUpper.Equals(acceptTypeList.Item(i)) Then
                rst = True
                Exit For
            End If
        Next

        Return rst
    End Function

#End Region

#Region "取得檔案大小"

    Function GetFileSize(ByVal fileLength As Integer) As String
        Dim rst As String = ""
        Dim tNum As Decimal = fileLength

        If tNum > 1024 Then
            tNum = tNum / 1024

            If tNum > 1024 Then
                tNum = tNum / 1024

                If tNum > 1024 Then

                    tNum = tNum / 1024

                    If tNum > 1024 Then

                        tNum = tNum / 1024
                        rst = Math.Round(tNum, 1).ToString + "TB"  '(暫時設定此系統最大容量為TB)
                    Else
                        rst = Math.Round(tNum, 1).ToString + "GB"
                    End If
                Else
                    rst = Math.Round(tNum, 1).ToString + "MB"
                End If
            Else
                rst = Math.Round(tNum, 1).ToString + "KB"
            End If
        Else
            rst = tNum.ToString + "Byte"
        End If

        rst = rst.Replace(".0", "")   '若檔案容量含有[.0]時,則將小數點拿掉

        '========== 20180718 Note:因Integer型別最大值對應到的大小為2GB，所以[檔案大小]最大值為2GB

        Return rst
    End Function

#End Region

    Function Get_F_GROUPID() As Integer
        Dim iF_GROUPID As Integer = 1
        'Dim vClassID_SEQ As String="TB_FILE_F_GROUPID_SEQ"/TB_CONTENT_F_GROUPID_SEQ
        Using tConn As SqlConnection = DbAccess.GetConnection()
            DbAccess.Open(tConn)
            Dim oTrans As SqlTransaction = DbAccess.BeginTrans(tConn)
            iF_GROUPID = DbAccess.GetAutoNum(tConn, oTrans, "TB_CONTENT_F_GROUPID_SEQ")  ' 忽略: TABLE_NAME, TABLE_PK 這二個參數
            DbAccess.CommitTrans(oTrans)
            DbAccess.CloseDbConn(tConn)
        End Using
        Return iF_GROUPID
    End Function

    Private Sub gv1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gv1.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                Dim drv As DataRowView = e.Row.DataItem
                'Dim secNo As Integer=Convert.ToInt32(drv("SEC_NO"))
                'Dim secNo_sys As Integer=secNo - 1
                Dim txtData As TextBox = e.Row.FindControl("txtData")
                Dim lblSEQ As Label = e.Row.FindControl("lblSEQ")
                Dim imgFUrl As WebControls.Image = e.Row.FindControl("imgFUrl")
                Dim lblFName As Label = e.Row.FindControl("lblFName")
                Dim divPic As HtmlControls.HtmlGenericControl = e.Row.FindControl("divPic")
                Dim rblPosition As RadioButtonList = e.Row.FindControl("rblPosition")
                Dim txtPicAlt As TextBox = e.Row.FindControl("txtPicAlt")

                txtData.Text = Convert.ToString(drv("SEC_CONTENT"))
                lblSEQ.Text = Convert.ToString(drv("SEQNO"))
                Dim picName As String = Convert.ToString(drv("SEC_PICTURE"))
                divPic.Visible = False
                If picName <> "" Then
                    divPic.Visible = True
                    imgFUrl.ImageUrl = ImageUrl_Path1 + picName
                    lblFName.Text = picName

                    txtPicAlt.Text = Convert.ToString(drv("SEC_PICTURE_ALT")) '有上傳圖檔才顯示圖檔提示文字
                End If
                Common.SetListItem(rblPosition, Convert.ToString(drv("ALIGN_TYPE")))
        End Select
    End Sub

    Private Sub gv2_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gv2.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                Dim drV As DataRowView = e.Row.DataItem

                Dim F_Name1 As String = Convert.ToString(drV("FILE_ORINAME"))
                Dim F_Name2 As String = Convert.ToString(drV("FILE_PHYNAME"))
                Dim F_Type As String = Convert.ToString(drV("FILE_TYPE"))
                Dim F_Size As String = Convert.ToString(drV("FILE_SIZE"))

                Dim lnkFName As LinkButton = CType(e.Row.FindControl("lnkFName"), LinkButton)
                Dim LinkButton As Label = CType(e.Row.FindControl("lblFName"), Label)
                Dim lblFExt As Label = CType(e.Row.FindControl("lblFExt"), Label)
                Dim lblFFileid As Label = CType(e.Row.FindControl("lblFFileid"), Label)
                F_Name1 = TIMS.ClearSQM(F_Name1)
                F_Type = TIMS.ClearSQM(F_Type)
                lnkFName.Text = F_Name1
                lnkFName.ToolTip = "[ 附檔名：" + F_Type.ToLower.Substring(1) + "  /  檔案大小：" + F_Size + " ]"
                F_Name2 = TIMS.ClearSQM(F_Name2)
                LinkButton.Text = F_Name2
                lblFExt.Text = F_Type.ToUpper
                lblFFileid.Text = Convert.ToString(drV("FILEID"))
        End Select

    End Sub
End Class