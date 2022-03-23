<%@ Page Language="VB" AutoEventWireup="false" CodeFile="v_log1_in.aspx.vb" Inherits="v_log1_in" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    
<script language="javascript" type="text/javascript">
<!--

function TABLE1_onclick() {

}

// -->
</script>
</head>

<body bgcolor="#99ccff" style="text-align: center;font-size: 12pt; background-color: #ccffff;" topmargin="3">


    
      <form id="form1" runat="server">
        <table style="width: 760px">
          <tr>
            <td colspan="3">
              <asp:Image ID="Image1" runat="server" Height="86px" ImageUrl="~/image/FAR_TOP2.JPG"
                Width="760px" /></td>
          </tr>
            <tr><td colspan="3" style ="text-align: right ">  
            <asp:Label ID="Label1" runat="server" Text="Label" BackColor="#FFC080" BorderColor="Black" BorderStyle="Solid" BorderWidth="1px" Font-Size="Medium" height="23px"></asp:Label>
            <asp:button id="Button1" runat="server" Width="50px" height="25px" Text="登出" ForeColor="Black" BorderColor="black" Font-Size="Medium"
                BackColor="#E0E0E0" BorderStyle="Solid" BorderWidth="1px"></asp:button>
        </td></tr>
          <tr>
            <td colspan="3">
              <br />
                  <asp:Label ID="lblMDB" runat="server" Font-Size="Large" ForeColor="Blue" Text="lblMDB" BackColor="#C0FFFF"></asp:Label><br />
                <strong>【1.<span style="font-size: 12pt; font-family: 細明體; mso-bidi-font-family: 'Times New Roman';
                    mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                    mso-bidi-language: AR-SA">電話錄音線路使用統計表</span>】</strong><br />
              <br />
              <table style="HEIGHT: 222px; font-family: Times New Roman; background-color: black;" height="222" cellSpacing="1" cellPadding="1" width="600" bgColor="#0080c0"
                border="0">
                <tr bgColor="#ffffff">
                  <td width="556" bgColor="#0080c0" bordercolor="#ff0033" style="height: 14px; background-color: #ccffff">
                    <p class="style5" style="MARGIN: 0px 4px; LINE-HEIGHT: 150%; text-align: center;" align="center"><font face="新細明體" size="2"><b><span style="font-size: 1.4em">&nbsp;<span style="font-size: 12pt;
                            font-family: Times New Roman">輸入查詢時間</span></span></b></font></p>
                  </td>
                </tr>
                <tr bgColor="#ffffff" style="font-size: 12pt; font-family: Times New Roman;">
                  <td vAlign="middle" align="left" background="#99ccff" bgcolor="#99ccff" style="height: 46px; background-color: #ccffff;" bordercolor="#ff0033">
                    <blockquote>
                      <p><font size="2">From:&nbsp;</font>
                        <asp:dropdownlist id="Year1" runat="server" Height="40px" Width="65px" BackColor="#C0C0FF"></asp:dropdownlist>Year
                        <asp:dropdownlist id="Month1" runat="server" Height="40px" Width="48px" BackColor="#C0C0FF"></asp:dropdownlist>Month
                        <asp:dropdownlist id="Day1" runat="server" Height="40px" Width="48px" BackColor="#C0C0FF"></asp:dropdownlist>Day
                        <asp:dropdownlist id="Hour1" runat="server" Width="48px" Height="40px" BackColor="#C0C0FF"></asp:dropdownlist>Hour&nbsp;
                        <asp:dropdownlist id="Min1" runat="server" Width="48px" Height="40px" BackColor="#C0C0FF"></asp:dropdownlist>Minute
                      </p>
                      <p><font size="2"><span style="font-size: 12pt; background-color: #ccffff;">To:</span> &nbsp;&nbsp; &nbsp;
                          <asp:dropdownlist id="Year2" runat="server" Height="40px" Width="65px" BackColor="#C0C0FF"></asp:dropdownlist>
Year</font>&nbsp;
                        <asp:dropdownlist id="Month2" runat="server" Height="40px" Width="48px" BackColor="#C0C0FF"></asp:dropdownlist>Month
                        <asp:dropdownlist id="Day2" runat="server" Height="40px" Width="48px" BackColor="#C0C0FF"></asp:dropdownlist>Day
                        <asp:dropdownlist id="Hour2" runat="server" Width="48px" Height="40px" BackColor="#C0C0FF"></asp:dropdownlist>Hour&nbsp;
                        <asp:dropdownlist id="Min2" runat="server" Width="48px" Height="40px" BackColor="#C0C0FF"></asp:dropdownlist>Minute
                      </p>
                      <p>
                        <asp:button id="butRun" runat="server" Width="56px" Text="確定" ForeColor="#0000C0" BorderColor="#00AEEF"
                          BackColor="#E0E0E0" BorderStyle="Solid"></asp:button>&nbsp;&nbsp;
                        <asp:button id="butReset" runat="server" Width="56px" Text="修改" ForeColor="#0000C0" BorderColor="#00AEEF"
                          BackColor="#E0E0E0" BorderStyle="Solid"></asp:button>
                        &nbsp;
              <asp:button id="butMenu" runat="server" Width="95px" Text="回主選單" ForeColor="#0000C0" BorderColor="#00AEEF"
                BackColor="#E0E0E0" BorderStyle="Solid"></asp:button></p>
                      <p style="background-color: #ccffff">
                        <asp:Label id="warning" runat="server" Width="500px" Height="19px" Visible="False" BackColor="#C0FFFF"
                          BorderColor="Transparent" ForeColor="Blue">Start Date error , Please reset.</asp:Label>&nbsp;</p>
                    </blockquote>
                  </td>
                  
                </tr>
              </table>
              <br />
              <br />
            </td>
          </tr>
          <tr style="font-family: Times New Roman">
            <td style="width: 100px">
            </td>
            <td style="width: 100px">
            </td>
            <td style="width: 100px">
            </td>
          </tr>
        </table>
        <br />
     </form>
  </body>

</html>
