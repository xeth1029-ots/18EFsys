<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_004_InsertPlan.aspx.vb" Inherits="WDAIIP.TC_01_004_InsertPlan" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>開班資料轉入</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="JavaScript">
        function fnEnable() {
            document.form1.change.disabled = false;
        }

        function wopen(url, name, width, height, k) {
            LeftPosition = (screen.width) ? (screen.width - width) / 2 : 0;
            TopPosition = (screen.availHeight) ? (screen.availHeight - height - 28) / 2 : 0;
            window.open(url, name, 'top=' + TopPosition + ',left=' + LeftPosition + ',width=' + width + ',height=' + height + ',resizable=1,scrollbars=' + k + ',status=0');
        }

        //function change_onclick() {
        //}
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;開班資料設定&gt;&gt;開班資料轉入</asp:Label>
                </td>
            </tr>
        </table>

        <table id="FrameTable3" cellspacing="0" cellpadding="0" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" width="20%">機構名稱							</td>
                            <td class="whitecol" colspan="3">&nbsp;<asp:Label ID="Name" runat="server"></asp:Label>
                                <input id="clsid" type="hidden" runat="server">
                                <input id="Re_ID" type="hidden" name="Re_ID" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">訓練計畫</td>
                            <td class="whitecol" width="30%">&nbsp;<asp:Label ID="Plan_Name" runat="server"></asp:Label></td>
                            <td class="bluecol" width="20%"><asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label></td>
                            <td class="whitecol" width="30%">
                                <font>&nbsp;[</font>
                                <asp:Label ID="TrainID" runat="server"></asp:Label>
                                <font>]</font>
                                <asp:Label ID="Train_Name" runat="server"></asp:Label>
                                <asp:Label ID="Job_Name" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%"><asp:Label ID="Labcjob" runat="server">通俗職類</asp:Label></td>
                            <td class="whitecol">
                                <font>&nbsp; [</font><asp:Label ID="cjobValue" runat="server"></asp:Label><font>]</font>
                                <asp:Label ID="txtCJOB_NAME" runat="server"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <br />
        <table id="search_tbl" class="font" border="1" cellspacing="0" cellpadding="0" width="100%" runat="server"></table>
        <table class="font" border="0" cellspacing="0" cellpadding="0" width="100%">
            <tr>
                <td>
                    <font color="#ff0000" size="2">
                        <br />※注意※<br />
                        當【班級申請】功能的《通俗職類》設定與【班別代碼】功能的《通俗職類》設定不同時，<br />
                        系統將以【班級代碼】功能的《通俗職類》為主，覆蓋掉【班級申請】功能的《通俗職類》。<br />
                    </font>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <input id="return_class" value="回開班設定" type="button" name="Submit" runat="server" class="asp_Export_M" />
                    <%--
					<input id="change" value="轉入" type="button" name="change" runat="server" onclick="return change_onclick()">
					<input id="change" value="轉入" type="button" name="change" runat="server" />
                    --%>
                    <input id="change" value="轉入" type="button" name="change" runat="server" class="asp_Export_M" />
                    <input id="back" value="回上一頁" type="button" name="back" runat="server" class="asp_Export_M" />
                </td>
            </tr>
        </table>
        <input id="isBlack" type="hidden" name="isBlack" runat="server" />
        <input id="orgname" type="hidden" name="orgname" runat="server" />
        <asp:HiddenField ID="Hid_RID1" runat="server" />
        <asp:HiddenField ID="Hid_ComIDNO" runat="server" />
    </form>
</body>
</html>