<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="RWB_01_001_edit.aspx.vb" Inherits="WDAIIP.RWB_01_001_edit" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>最新消息</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker2.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <%--<script type="text/javascript" language="javascript">
        //決定date-picker元件使用的是西元年or民國年，by:20181019
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);
    </script>--%>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>

</head>
<body>
    <form id="form1" method="post" runat="server" enctype="multipart/form-data">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;報名網維護&gt;&gt;最新消息</asp:Label>
                </td>
            </tr>
        </table>
        <br>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td width="20%" class="bluecol_need">類別：</td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlType" runat="server" Width="30%">
                        <asp:ListItem Value="1" Selected="True">焦點消息</asp:ListItem>
                        <asp:ListItem Value="2">計畫公告</asp:ListItem>
                        <asp:ListItem Value="3">成果集錦</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">上稿日期：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtCDATE1" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox></td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">上架日期：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtSDATE1" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                    <span id="span1" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txtSDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span2" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= txtSDATE1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                    <asp:DropDownList ID="ddlC_SDATE_hh1" runat="server"></asp:DropDownList>時
                    <asp:DropDownList ID="ddlC_SDATE_mm1" runat="server"></asp:DropDownList>分
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">標題：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtTitle" Width="80%" runat="server" placeholder="請輸入標題內容" MaxLength="300"></asp:TextBox></td>
            </tr>
            <tr id="trCSORT1" runat="server">
                <td width="20%" class="bluecol">排序序號：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtCSORT1" Width="20%" runat="server" placeholder="請輸入序號" MaxLength="5"></asp:TextBox>
                    <asp:Label ID="lab_msg_CSORT1" runat="server" ForeColor="Red" Text="(由小到大排序)"></asp:Label>
                </td>
            </tr>
            <tr id="trC_URL" runat="server">
                <td width="20%" class="bluecol_need">宣導影片網址：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtC_URL" Width="88%" runat="server" placeholder="請輸入 宣導影片網址" MaxLength="4000"></asp:TextBox>
                </td>
            </tr>
            <tr id="trC_CONTENT1" runat="server">
                <td width="20%" class="bluecol_need">宣導連結：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtC_CONTENT1" Width="88%" runat="server" placeholder="請輸入 宣導連結" MaxLength="4000"></asp:TextBox>
                </td>
            </tr>
            <tr id="trLINKURL1" runat="server">
                <td width="20%" class="bluecol"><font color="red">#[url1]</font>／連結網址：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtLINKURL1" Width="88%" runat="server" placeholder="請輸入 連結網址1" MaxLength="4000"></asp:TextBox>
                </td>
            </tr>
            <tr id="trLINKURL2" runat="server">
                <td width="20%" class="bluecol"><font color="red">#[url2]</font>／連結網址：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtLINKURL2" Width="88%" runat="server" placeholder="請輸入 連結網址2" MaxLength="4000"></asp:TextBox>
                </td>
            </tr>
             <tr id="trLINKURL3" runat="server">
                <td width="20%" class="bluecol"><font color="red">#[url3]</font>／連結網址：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtLINKURL3" Width="88%" runat="server" placeholder="請輸入 連結網址3" MaxLength="4000"></asp:TextBox>
                </td>
            </tr>
             <tr id="trLINKURL4" runat="server">
                <td width="20%" class="bluecol"><font color="red">#[url4]</font>／連結網址：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtLINKURL4" Width="88%" runat="server" placeholder="請輸入 連結網址4" MaxLength="4000"></asp:TextBox>
                </td>
            </tr>
             <tr id="trLINKURL5" runat="server">
                <td width="20%" class="bluecol"><font color="red">#[url5]</font>／連結網址：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtLINKURL5" Width="88%" runat="server" placeholder="請輸入 連結網址5" MaxLength="4000"></asp:TextBox>
                </td>
            </tr>

            <tr>
                <td colspan="2">
                    <asp:GridView ID="gv1" runat="server" AutoGenerateColumns="False" Width="100%" ShowHeader="False" BorderStyle="None">
                        <Columns>
                            <asp:TemplateField HeaderText="段落內容">
                                <ItemTemplate>
                                    <table class="table_nw" width="100%" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="bluecol" width="20%">內容<asp:Label ID="lblT1" runat="server" Text='<%# Bind("item") %>'></asp:Label>：</td>
                                            <td class="whitecol" width="80%">
                                                <asp:TextBox ID="txtData" Width="80%" runat="server" placeholder="請輸入內容" TextMode="MultiLine" Rows="16"></asp:TextBox>
                                                <asp:Label ID="lblSEQ" runat="server" Text="" Visible="false"></asp:Label>
                                                <asp:Label ID="lblSecNo" runat="server" Text='<%# Bind("id") %>' Visible="false"></asp:Label>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol" width="20%">
                                                <asp:Label ID="lblT2" runat="server" Text='<%# Bind("item") %>'></asp:Label>圖檔：</td>
                                            <td class="whitecol" width="80%">
                                                <asp:FileUpload ID="fu1" runat="server" Width="60%" />
                                                <asp:Button ID="bt_upPic" runat="server" Text="上傳" CssClass="asp_button_M" AuthType="UPPIC" OnClick="bt_upPic_Click"></asp:Button>
                                                <div id="divPic" runat="server" visible="false">
                                                    <br />
                                                    &nbsp;&nbsp;&nbsp;
                                                    <div style="max-width: 600px; max-height: 400px; overflow-x: auto; overflow-y: auto;">
                                                        <asp:Image ID="imgFUrl" runat="server" Visible="true" AlternateText="(圖不存在)" />
                                                    </div>
                                                    &nbsp;<asp:Label ID="lblFName" runat="server" Visible="false"></asp:Label>
                                                    &nbsp;<asp:Button ID="bt_delPic" runat="server" Text="刪除" CssClass="asp_button_M" AuthType="DELPIC" OnClick="bt_delPic_Click"></asp:Button>
                                                </div>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">圖檔提示文字：</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="txtPicAlt" Width="80%" MaxLength="40" runat="server"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol" width="20%">圖檔位置：</td>
                                            <td class="whitecol" width="80%">
                                                <asp:RadioButtonList ID="rblPosition" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                                    <asp:ListItem Value="L" Selected="True">圖靠左</asp:ListItem>
                                                    <asp:ListItem Value="R">圖靠右</asp:ListItem>
                                                </asp:RadioButtonList>
                                            </td>
                                        </tr>
                                    </table>
                                </ItemTemplate>
                                <ItemStyle HorizontalAlign="Left" VerticalAlign="Top" Width="100%" BorderStyle="None" />
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">停用日期：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtEDATE1" Width="20%" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="span3" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txtEDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span4" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= txtEDATE1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                    <asp:DropDownList ID="ddlC_EDATE_hh1" runat="server"></asp:DropDownList>時
                    <asp:DropDownList ID="ddlC_EDATE_mm1" runat="server"></asp:DropDownList>分
                </td>
            </tr>
            <tr id="trUPFILE1" runat="server">
                <td width="20%" class="bluecol">附件檔案：</td>
                <td class="whitecol">
                    <br />
                    <asp:FileUpload ID="fu2" runat="server" Width="60%" />
                    <asp:Button ID="bt_upfile" runat="server" Text="上傳" CssClass="asp_button_M" AuthType="UPFILE"></asp:Button>
                    <div id="divFile" runat="server" visible="false">
                        <br />
                        <asp:GridView ID="gv2" runat="server" AutoGenerateColumns="False" Width="80%" ShowHeader="False" BorderStyle="None">
                            <Columns>
                                <asp:TemplateField HeaderText="附件內容">
                                    <ItemTemplate>
                                        <asp:LinkButton ID="lnkFName" runat="server" Text="(附件檔案)" OnClick="lnkFName_Click"></asp:LinkButton>
                                        &nbsp;<asp:Label ID="lblFName" runat="server" Visible="false"></asp:Label>
                                        &nbsp;<asp:Label ID="lblFExt" runat="server" Visible="false"></asp:Label>
                                        &nbsp;<asp:Label ID="lblFFileid" runat="server" Visible="false"></asp:Label>
                                        &nbsp;<asp:Button ID="bt_delFile" runat="server" Text="刪除" CssClass="asp_button_M" AuthType="DELFILE" OnClick="bt_delFile_Click"></asp:Button>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle" Width="100%" BorderStyle="None" />
                                </asp:TemplateField>
                            </Columns>
                        </asp:GridView>
                    </div>
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol" align="center">
                    <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_button_S" AuthType="SAVE"></asp:Button>
                    <asp:Button ID="bt_cancle" runat="server" Text="取消" CausesValidation="False" CssClass="asp_button_S" AuthType="CANCLE"></asp:Button>&nbsp;
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
            </tr>
        </table>
        <input type="hidden" runat="server" id="hid_V_SEQNO" />
        <input type="hidden" runat="server" id="hid_f_grp" />
    </form>
</body>
</html>
