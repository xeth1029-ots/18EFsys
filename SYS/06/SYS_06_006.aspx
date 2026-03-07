<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_06_006.aspx.vb" Inherits="WDAIIP.SYS_06_006" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>個資安全密碼保護設定</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        var hexcase = 0; /* hex output format. 0 - lowercase; 1 - uppercase */
        var chrsz = 8; /* bits per input character. 8 - ASCII; 16 - Unicode */
        /* The main function to calculate message digest  */
        function hex_sha1(s) {
            return binb2hex(core_sha1(AlignSHA1(s)));
        }

        /* Calculate the SHA-1 of an array of big-endian words, and a bit length */
        function core_sha1(blockArray) {
            var x = blockArray; // append padding
            var w = Array(80);
            var a = 1732584193;
            var b = -271733879;
            var c = -1732584194;
            var d = 271733878;
            var e = -1009589776;
            for (var i = 0; i < x.length; i += 16) // 每次处理512位 16*32
            {
                var olda = a;
                var oldb = b;
                var oldc = c;
                var oldd = d;
                var olde = e;
                for (var j = 0; j < 80; j++) // 对每个512位进行80步操作
                {
                    if (j < 16)
                        w[j] = x[i + j];
                    else
                        w[j] = rol(w[j - 3] ^ w[j - 8] ^ w[j - 14] ^ w[j - 16], 1);
                    var t = safe_add(safe_add(rol(a, 5), sha1_ft(j, b, c, d)), safe_add(safe_add(e, w[j]), sha1_kt(j)));
                    e = d;
                    d = c;
                    c = rol(b, 30);
                    b = a;
                    a = t;
                }
                a = safe_add(a, olda);
                b = safe_add(b, oldb);
                c = safe_add(c, oldc);
                d = safe_add(d, oldd);
                e = safe_add(e, olde);
            }
            return new Array(a, b, c, d, e);
        }

        /*
         * Perform the appropriate triplet combination function for the current
         * iteration
         * 返回对应F函数的值
         */
        function sha1_ft(t, b, c, d) {
            if (t < 20)
                return (b & c) | ((~b) & d);

            if (t < 40)
                return b ^ c ^ d;

            if (t < 60)
                return (b & c) | (b & d) | (c & d);

            return b ^ c ^ d; // t<80
        }

        /*
         * Determine the appropriate additive constant for the current iteration
         * 返回对应的Kt值
         */
        function sha1_kt(t) {
            return (t < 20) ? 1518500249 : (t < 40) ? 1859775393 : (t < 60) ? -1894007588 : -899497514;
        }

        /*
         * Add integers, wrapping at 2^32. This uses 16-bit operations internally
         * to work around bugs in some JS interpreters.
         * 将32位数拆成高16位和低16位分别进行相加，从而实现 MOD 2^32 的加法
         */
        function safe_add(x, y) {
            var lsw = (x & 0xFFFF) + (y & 0xFFFF);
            var msw = (x >> 16) + (y >> 16) + (lsw >> 16);
            return (msw << 16) | (lsw & 0xFFFF);
        }

        /*
         * Bitwise rotate a 32-bit number to the left.
         * 32位二进制数循环左移
         */
        function rol(num, cnt) {
            return (num << cnt) | (num >>> (32 - cnt));
        }

        /*
        The standard SHA1 needs the input string to fit into a block
        This function align the input string to meet the requirement
        */
        function AlignSHA1(str) {
            var nblk = ((str.length + 8) >> 6) + 1, blks = new Array(nblk * 16);
            for (var i = 0; i < nblk * 16; i++)
                blks[i] = 0;
            for (i = 0; i < str.length; i++)
                blks[i >> 2] |= str.charCodeAt(i) << (24 - (i & 3) * 8);
            blks[i >> 2] |= 0x80 << (24 - (i & 3) * 8);
            blks[nblk * 16 - 1] = str.length * 8;
            return blks;
        }


        /*Convert an array of big-endian words to a hex string.*/
        function binb2hex(binarray) {
            var hex_tab = hexcase ? "0123456789ABCDEF" : "0123456789abcdef";
            var str = "";
            for (var i = 0; i < binarray.length * 4; i++) {
                str += hex_tab.charAt((binarray[i >> 2] >> ((3 - i % 4) * 8 + 4)) & 0xF) +
                hex_tab.charAt((binarray[i >> 2] >> ((3 - i % 4) * 8)) & 0xF);
            }
            return str;
        }

        /* Perform a simple self-test to see if the VM is working */
        function sha1_vm_test() {
            return hex_sha1("abc") == "a9993e364706816aba3e25717850c26c9cd0d89d";
        }


        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length; //長度
            var myallcheck = document.getElementById(obj + '_' + 0); //第1個

            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        var mycheck = document.getElementById(obj + '_' + i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <div style="position: absolute; top: -333px"><input type="text" title="Chaff for Chrome Smart Lock" /><input type="password" title="Chaff for Chrome Smart Lock" /></div>
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;個資安全密碼保護設定</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">

            <tr>
                <td>
                    <table id="Table3" class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 20%">年度
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="list_Years" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">轄區
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="list_DistID" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                                &nbsp;
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">計畫代碼
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="list_PlanID" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">密碼
                            </td>
                            <td class="whitecol">
                                <asp:TextBox Style="z-index: 0" ID="tPXssXArd1" runat="server" MaxLength="30" TextMode="Password" Width="40%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">重key密碼
                            </td>
                            <td class="whitecol">
                                <asp:TextBox Style="z-index: 0" ID="tPXssXArd2" runat="server" MaxLength="30" TextMode="Password" Width="40%"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="btnQuery" runat="server" Text="重設密碼" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <%--<table id="DataGridTable1" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
					<tr>
						<td align="center">
							<asp:DataGrid Style="z-index: 0" ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font">
								<AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
								<ItemStyle BackColor="#EBF3FE"></ItemStyle>
								<HeaderStyle HorizontalAlign="Center" BackColor="#96B5E3"></HeaderStyle>
								<Columns>
									<asp:BoundColumn HeaderText="編號">
										<HeaderStyle Width="30px"></HeaderStyle>
										<ItemStyle HorizontalAlign="Center"></ItemStyle>
									</asp:BoundColumn>
									<asp:BoundColumn DataField="ACCOUNT" HeaderText="使用者帳號"></asp:BoundColumn>
									<asp:BoundColumn DataField="CName" HeaderText="姓名"></asp:BoundColumn>
									<asp:BoundColumn DataField="WorkDate" HeaderText="作業日期"></asp:BoundColumn>
								</Columns>
								<PagerStyle Visible="False"></PagerStyle>
							</asp:DataGrid>
						</td>
					</tr>
					<tr>
						<td style="height: 31px" align="center">
							<uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
						</td>
					</tr>
				</table>--%>
                </td>
            </tr>
        </table>
        <%--
				<TR>
					<TD>
					</TD>
				</TR>
        --%>
        <input id="hidPXwSd" type="hidden" name="hidPXwSd" runat="server">
    </form>
</body>
</html>
