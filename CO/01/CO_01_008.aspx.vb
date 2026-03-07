Imports System.IO
Imports OfficeOpenXml

Public Class CO_01_008
    Inherits AuthBasePage 'System.Web.UI.Page

    '排程[Co_OrgScoring] Co_OrgScoring.exe.config '排程[-xx-CO_ORGSCORING] /'ORG_SCORING2
    'Lab_SUSPENDED_msg1 'Const cst_SUSPENDED_msgFM1 As String = "此單位因有{0}班停班經認列屬「不可抗力因素」，將不列入核定總班數計算。"
    ReadOnly ss_Search1 As String = "CO_01_008_Search1"
    ReadOnly cst_PageSort As String = "PageSort"
    Const Cst_DG_COL_DISTID As Integer = 1
    Const Cst_DG_COL_ORGNAME As Integer = 2
    Const Cst_DG_COL_COMIDNO As Integer = 3
    Const Cst_DG_COL_RLEVEL_2 As Integer = 4

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            '(加強操作便利性)
            cCreate1()
        End If
        If sm.UserInfo.DistID <> "000" Then
            Common.SetListItem(ddlDISTID, sm.UserInfo.DistID)
            ddlDISTID.Enabled = False
        End If
    End Sub

    Sub KeepSearch1()
        '(加強操作便利性)
        Dim s_kpSearch1 As String = ""
        TIMS.SetMyValue(s_kpSearch1, "ddlDISTID", TIMS.GetListValue(ddlDISTID))
        TIMS.SetMyValue(s_kpSearch1, "ddlSCORING", TIMS.GetListValue(ddlSCORING))
        TIMS.SetMyValue(s_kpSearch1, "OrgName", OrgName.Text)
        TIMS.SetMyValue(s_kpSearch1, "COMIDNO", COMIDNO.Text)
        'TIMS.SetMyValue(s_kpSearch1, "OrgPlanKind", TIMS.GetListValue(OrgPlanKind))
        'TIMS.SetMyValue(s_kpSearch1, "OrgKindList", TIMS.GetListValue(OrgKindList))
        Session(ss_Search1) = s_kpSearch1
    End Sub

    Sub UseKeepSearch1()
        '(加強操作便利性)
        If Session(ss_Search1) Is Nothing Then Return
        Dim s_kpSearch1 As String = Session(ss_Search1)
        If s_kpSearch1 = "" Then Return
        Session(ss_Search1) = Nothing
        Common.SetListItem(ddlDISTID, TIMS.GetMyValue(s_kpSearch1, "ddlDISTID"))
        Common.SetListItem(ddlSCORING, TIMS.GetMyValue(s_kpSearch1, "ddlSCORING"))
        OrgName.Text = TIMS.GetMyValue(s_kpSearch1, "OrgName")
        COMIDNO.Text = TIMS.GetMyValue(s_kpSearch1, "COMIDNO")
        'Common.SetListItem(OrgPlanKind, TIMS.GetMyValue(s_kpSearch1, "OrgPlanKind"))
        'Common.SetListItem(OrgKindList, TIMS.GetMyValue(s_kpSearch1, "OrgKindList"))
        Call sSearch1()
    End Sub

    Sub cCreate1()
        divSch1.Visible = True
        msg1.Text = ""
        PageControler1.Visible = False
        '評核版本 'ddlSENDVER = Get_SENDVER_TS(ddlSENDVER) '評核結果 'ddlRESULT = Get_RESULT_TS(ddlRESULT)

        ddlDISTID = TIMS.Get_DistID(ddlDISTID, Nothing, objconn)
        If (ddlDISTID.Items.FindByValue("000") IsNot Nothing) Then ddlDISTID.Items.Remove(ddlDISTID.Items.FindByValue("000"))
        Common.SetListItem(ddlDISTID, sm.UserInfo.DistID)

        ddlSCORING = TIMS.Get_ddlSCORING(ddlSCORING, objconn)
        'SYEARlist = TIMS.GetSyear(SYEARlist) 'Common.SetListItem(SYEARlist, sm.UserInfo.Years)

        '依登入者機構判斷計畫種類 '依登入者 LID 判斷是否可自由輸入
        If sm.UserInfo.LID = 2 Then '委訓單位動作
            Dim droo As DataRow = TIMS.Get_ORGINFOdr(sm.UserInfo.OrgID, objconn)
            OrgName.Text = Convert.ToString(droo("OrgName"))
            COMIDNO.Text = Convert.ToString(droo("ComIDNO"))
            'OrgName.Enabled = False 'COMIDNO.Enabled = False
        End If

        '階層代碼 0:署 1:中心 2:委訓
        If sm.UserInfo.LID <> 0 Then
            ddlDISTID.Enabled = False '1/2
            If sm.UserInfo.LID <> 1 Then
                OrgName.Enabled = False '2
                COMIDNO.Enabled = False '2
            End If
        End If

        '登入年度轉民國年份 'Years.Value = sm.UserInfo.Years - 1911

        '選擇清除工作 'SelectValue.Value = ""
        DataGridTable.Visible = False
        Call UseKeepSearch1()
    End Sub

    Protected Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Call sSearch1()
    End Sub

    Sub sSearch1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        PageControler1.Visible = False
        DataGridTable.Visible = False
        msg1.Text = "查無資料"

        Dim vDISTID As String = TIMS.ClearSQM(ddlDISTID.SelectedValue)
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        'Dim vYEARS As String = TIMS.ClearSQM(SYEARlist.SelectedValue)
        'Dim vHALFYEAR As String = TIMS.ClearSQM(halfYear.SelectedValue) '1:上年度 /2:下年度

        Dim vORGNAME As String = TIMS.ClearSQM(OrgName.Text)
        Dim vCOMIDNO As String = TIMS.ClearSQM(COMIDNO.Text)
        'Dim vORGKIND2 As String = TIMS.ClearSQM(OrgPlanKind.SelectedValue) '計畫
        'Dim vORGKIND As String = TIMS.ClearSQM(OrgKindList.SelectedValue) '機構別

        Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        'Dim flag_CanNoDistIDValue As Boolean
        Dim eErrMsg1 As String = ""
        If Not flagS1 AndAlso vDISTID = "" Then eErrMsg1 &= "請選擇分署" & vbCrLf

        'If vYEARS = "" Then eErrMsg1 &= "請選擇年度" & vbCrLf
        If vSCORINGID = "" Then eErrMsg1 &= "請選擇 審查計分區間" & vbCrLf
        If eErrMsg1 <> "" Then
            Common.MessageBox(Me, eErrMsg1)
            Exit Sub
        End If

        Call KeepSearch1()

        Dim parms As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"SCORINGID", vSCORINGID}}
        If vDISTID <> "" Then parms.Add("DISTID", vDISTID)
        If vORGNAME <> "" Then parms.Add("ORGNAME", vORGNAME)
        If vCOMIDNO <> "" Then parms.Add("COMIDNO", vCOMIDNO)

        Dim sSql As String = "
SELECT a.OSID2 ,a.OrgID,a.TPLANID,a.DISTID
,CONCAT(dbo.FN_CYEAR2(a.YEARS) ,'年',a.MONTHS,'月'
 ,'(',dbo.FN_CYEAR2(a.YEARS1) ,'年',case when a.HALFYEAR1=1 then '上半年' else '下半年' end ,'~'
 ,dbo.FN_CYEAR2(a.YEARS2) ,'年',case when a.HALFYEAR2=1 then '上半年' else '下半年' end ,')') SCORING_N
,CONCAT(a.YEARS ,'-',a.MONTHS,'-',a.YEARS1 ,'-',a.HALFYEAR1,'-',a.YEARS2 ,'-',a.HALFYEAR2) SCORINGID  
,ISNULL(a.BRANCHPNT,0) BRANCHPNT,ISNULL(a.SUBTOTAL,0) SUBTOTAL
,a.IMPSCORE_1,a.IMPLEVEL_1,a.IMODIFYDATE,a.IMODIFYACCT
,a.RLEVEL_2,oo.ORGNAME,oo.COMIDNO,kd.DISTNAME3
,k1.NAME ORGKIND_N,v1.VNAME SENDVER_N,v2.VNAME RESULT_N
FROM dbo.ORG_SCORING2 a WITH(NOLOCK)
JOIN dbo.ORG_ORGINFO oo WITH(NOLOCK) ON oo.OrgID = a.OrgID
JOIN dbo.V_DISTRICT kd WITH(NOLOCK) ON kd.DISTID = a.DISTID COLLATE Chinese_Taiwan_Stroke_CS_AS
LEFT JOIN KEY_ORGTYPE k1 WITH(NOLOCK) ON k1.ORGTYPEID = oo.ORGKIND
LEFT JOIN dbo.ORG_TTQS2 b On concat(b.ORGID,'x',b.COMIDNO,b.TPLANID,b.DISTID,b.YEARS,b.MONTHS)=concat(a.ORGID,'x',a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS,a.MONTHS)
LEFT JOIN dbo.V_SENDVER v1 On v1.VID=b.SENDVER COLLATE Chinese_Taiwan_Stroke_CS_AS
LEFT JOIN dbo.V_RESULT v2 On v2.VID=b.RESULT COLLATE Chinese_Taiwan_Stroke_CS_AS AND v2.VID<='4'
"
        sSql &= " WHERE a.FIRSTCHK='Y'" & vbCrLf '(初審通過)
        '審查計分表-調修1：審查計分表(初審)、(複審)查詢清單顯示邏輯調整
        sSql &= " AND dbo.FN_GET_CLASSCNT2B(a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS1,a.YEARS2)>0" & vbCrLf
        sSql &= " AND a.TPLANID=@TPLANID" & vbCrLf
        sSql &= " AND CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)=@SCORINGID" & vbCrLf
        If vDISTID <> "" Then sSql &= " AND a.DISTID=@DISTID" & vbCrLf
        If vORGNAME <> "" Then sSql &= String.Concat(" AND oo.ORGNAME LIKE '%'+@ORGNAME+'%'")
        If vCOMIDNO <> "" Then sSql &= " AND oo.COMIDNO=@COMIDNO" & vbCrLf

        sSql &= " ORDER BY a.DISTID,oo.ORGNAME" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, parms)

        'If TIMS.sUtl_ChkTest() Then
        '    TIMS.writeLog(Me, String.Concat("--", vbCrLf, TIMS.GetMyValue5(parms)))
        '    TIMS.writeLog(Me, String.Concat("--##CO_01_008.aspx , sSql:", vbCrLf, sSql))
        'End If

        If TIMS.dtNODATA(dt) Then Return

        PageControler1.Visible = True
        DataGridTable.Visible = True
        msg1.Text = ""

        If ViewState(cst_PageSort) = "" Then ViewState(cst_PageSort) = "DISTID"
        PageControler1.PageDataTable = dt
        PageControler1.Sort = ViewState(cst_PageSort)
        PageControler1.ControlerLoad()
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Dim dg1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.Header
                If ViewState(cst_PageSort) <> "" Then
                    Dim mysort As New System.Web.UI.WebControls.Image
                    Dim i_Cell As Integer = -1
                    Dim str_Sort As String = Convert.ToString(ViewState(cst_PageSort))
                    Call ACT_ImageUrl_UD(mysort, i_Cell, str_Sort)
                    ViewState(cst_PageSort) = str_Sort
                    If i_Cell <> -1 Then e.Item.Cells(i_Cell).Controls.Add(mysort)
                End If

            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                'Dim lbtView As LinkButton = e.Item.FindControl("lbtView")
                Dim labORGNAME As Label = e.Item.FindControl("labORGNAME")
                Dim lRLEVEL_2X As Label = e.Item.FindControl("lRLEVEL_2X") '複審等級/複審<br>等級"
                labORGNAME.Text = Convert.ToString(drv("ORGNAME"))
                lRLEVEL_2X.Text = Convert.ToString(drv("RLEVEL_2"))
                'Dim sCmdArg As String = "" 'TIMS.SetMyValue(sCmdArg, "OSID2", Convert.ToString(drv("OSID2"))) 'lbtView.CommandArgument = sCmdArg

        End Select
    End Sub

    ''' <summary>依目前排序 顯示正確的排序圖型</summary>
    ''' <param name="sortVal"></param>
    ''' <returns></returns>
    Function GET_ImageUrl_UD(ByRef sortVal As String, ByRef str_PageSort As String) As String
        Return If(str_PageSort.Equals(sortVal), "../../images/SortUp.gif", "../../images/SortDown.gif")
    End Function

    Sub ACT_ImageUrl_UD(ByRef mysort As System.Web.UI.WebControls.Image, ByRef i_Cell As Integer, ByRef str_PageSort As String)
        Select Case str_PageSort
            Case "DISTID", "DISTID DESC"
                i_Cell = Cst_DG_COL_DISTID '1
                mysort.ImageUrl = GET_ImageUrl_UD("DISTID", str_PageSort)
            Case "ORGNAME", "ORGNAME DESC"
                i_Cell = Cst_DG_COL_ORGNAME '2
                mysort.ImageUrl = GET_ImageUrl_UD("ORGNAME", str_PageSort)
            Case "COMIDNO", "COMIDNO DESC"
                i_Cell = Cst_DG_COL_COMIDNO '3
                mysort.ImageUrl = GET_ImageUrl_UD("COMIDNO", str_PageSort)
            Case "RLEVEL_2", "RLEVEL_2 DESC"
                i_Cell = Cst_DG_COL_RLEVEL_2 '4
                mysort.ImageUrl = GET_ImageUrl_UD("RLEVEL_2", str_PageSort)
        End Select
    End Sub

    Private Sub DataGrid1_SortCommand(source As Object, e As DataGridSortCommandEventArgs) Handles DataGrid1.SortCommand
        ViewState(cst_PageSort) = String.Concat(e.SortExpression, If(ViewState(cst_PageSort) <> e.SortExpression, "", " DESC"))
        PageControler1.Sort = Me.ViewState(cst_PageSort)
        Call sSearch1()
    End Sub

End Class
