<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="OB_01_001_add.aspx.vb"
    Inherits="WDAIIP.OB_01_001_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>OB_01_001_add</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="JavaScript">

        function set_planname1(obj) {
            //debugger;

            var TPlanID = document.getElementById('TPlanID');
            var PlanName = document.getElementById('PlanName');

            if (obj == 'rb1') {
                TPlanID.style.display = 'inline';
                PlanName.style.display = 'none';
            }

            if (obj == 'rb2') {
                TPlanID.style.display = 'none';
                PlanName.style.display = 'inline';
            }
        }
    </script>
    <style type="text/css">
        .style1
        {
            height: 22px;
        }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
										<FONT face="新細明體">首頁&gt;&gt;委外訓練管理&gt;&gt;<font color="#990000">委外訓練資料查詢</font></FONT>
                            </asp:Label><font color="#000000">(<font face="新細明體"><font color="#ff0000">*</font>為必填欄位</font>)</font>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <table class="table_sch" id="TableLay2" cellspacing="1" cellpadding="1">
                    <tr>
                        <td width="100" class="bluecol_need">
                            年度
                        </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="ddlyears" runat="server">
                            </asp:DropDownList>
                        </td>
                        <td width="100" class="bluecol_need">
                            序號
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="txttsn" runat="server" Enabled="False"></asp:TextBox><font color="#ff0000">(系統自動產生)</font>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            採購案類型
                        </td>
                        <td colspan="3" class="whitecol">
                            <asp:RadioButton ID="rb1" runat="server" GroupName="Type1" Text="訓練案"></asp:RadioButton>
                            <asp:RadioButton ID="rb2" runat="server" GroupName="Type1" Text="非訓練案"></asp:RadioButton>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need">
                            訓練計畫名稱
                        </td>
                        <td colspan="3" class="whitecol">
                            <asp:DropDownList ID="TPlanID" runat="server">
                            </asp:DropDownList>
                            <asp:TextBox ID="PlanName" runat="server" MaxLength="20"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need">
                            標案名稱
                        </td>
                        <td colspan="3" class="whitecol">
                            <asp:TextBox ID="TenderCName" runat="server" MaxLength="20" Width="400px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            標案英文名稱
                        </td>
                        <td colspan="3" class="whitecol">
                            <asp:TextBox ID="TenderEName" runat="server" MaxLength="100" Width="400px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            主辦單位
                        </td>
                        <td colspan="3" class="whitecol">
                            <asp:TextBox ID="Sponsor" runat="server" MaxLength="20" Width="400px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need">
                            投標日期
                        </td>
                        <td colspan="3" class="whitecol">
                            <asp:TextBox ID="TenderSDate" MaxLength="10" Width="80" runat="server"></asp:TextBox>
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= TenderSDate.ClientId %>','','','CY/MM/DD');"
                                alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            &nbsp;～&nbsp;
                            <asp:TextBox ID="TenderEDate" MaxLength="10" Width="80" runat="server"></asp:TextBox>
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= TenderEDate.ClientId %>','','','CY/MM/DD');"
                                alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            <asp:TextBox ID="TenderEDate2" MaxLength="5" Width="40" runat="server">17:00</asp:TextBox>
                            (時間格式為00:00~23:59)
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need">
                            評選日期
                        </td>
                        <td colspan="3" class="whitecol">
                            <asp:TextBox ID="ReviewDate" MaxLength="10" Width="80" runat="server"></asp:TextBox>
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= ReviewDate.ClientId %>','','','CY/MM/DD');"
                                alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            決標日期
                        </td>
                        <td colspan="3" class="whitecol">
                            <asp:TextBox ID="ResolutionDate" MaxLength="10" Width="80" runat="server"></asp:TextBox>
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= ResolutionDate.ClientId %>','','','CY/MM/DD');"
                                alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <p align="center">
                    <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button><font
                        face="新細明體">&nbsp;</font>
                    <asp:Button ID="btnReturn" runat="server" Text="回上一頁" CssClass="asp_button_S"></asp:Button></p>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
