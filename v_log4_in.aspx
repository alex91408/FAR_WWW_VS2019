<%@ Page Language="VB" AutoEventWireup="false" CodeFile="v_log4_in.aspx.vb" Inherits="v_log4_in" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    
</head>

  <body style="text-align: center; background-color: #ccffff;" bgcolor="#99ccff" topmargin="3">

    <form id="Form2" method="post" runat="server">
      <table style="width: 760px">
        <tr>
          <td colspan="3" style="text-align: center">
            <asp:Image ID="Image1" runat="server" Height="86px" ImageUrl="~/image/FAR_TOP2.JPG"
              Width="760px" /></td>
        </tr>
        <tr><td colspan="3" style ="text-align: right ">  
            <asp:Label ID="Label1" runat="server" Text="Label" BackColor="#FFC080" BorderColor="Black" BorderStyle="Solid" BorderWidth="1px" Font-Size="Medium" height="23px"></asp:Label>
            <asp:button id="Button1" runat="server" Width="50px" height="25px" Text="登出" ForeColor="Black" BorderColor="black" Font-Size="Medium"
                BackColor="#E0E0E0" BorderStyle="Solid" BorderWidth="1px"></asp:button>
        </td></tr>
        <tr>
          <td colspan="3" style="text-align: center">
            <br />
<br />
                  <asp:Label ID="lblMDB" runat="server" Font-Size="Large" ForeColor="Blue" Text="lblMDB" BackColor="#C0FFFF"></asp:Label><br />
              <strong>【4.<span style="font-size: 12pt; font-family: 細明體; mso-bidi-font-family: 'Times New Roman';
                  mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                  mso-bidi-language: AR-SA">各分行電話錄音統計表</span>】<br />
              </strong>

                <TABLE id="Table1" style="HEIGHT: 222px; background-color: black; text-align: center;" height="222" cellSpacing="1" cellPadding="1" width="600"
                  bgColor="#0080c0" border="0">
                  <TR bgColor="#ffffff">
                    <TD bgColor="#0080c0" style="height: 22px; width: 567px; background-color: #ccffff;">
                      <P class="style5" style="MARGIN: 0px 4px; LINE-HEIGHT: 150%; text-align: center;" align="center"><FONT face="新細明體" size="2"><B>&nbsp;<span style="font-size: 12pt; font-family: Times New Roman">
                          輸入查詢時間</span></B></FONT></P>
                    </TD>
                  </TR>
                  <TR bgColor="#ffffff" style="font-size: 12pt; font-family: Times New Roman">
                    <TD vAlign="middle" align="left" style="height: 115px; width: 567px; background-color: #ccffff;" bgcolor="#99ccff">
                      <BLOCKQUOTE>
                        <P>
<span style="font-size: 10pt">From:&nbsp;</span>
<asp:DropDownList ID="Year1" runat="server" Height="40px" Width="65px" BackColor="#C0C0FF">
</asp:DropDownList>Year
<asp:DropDownList ID="Month1" runat="server" Height="40px" Width="48px" BackColor="#C0C0FF">
</asp:DropDownList>Month
<asp:DropDownList ID="Day1" runat="server" Height="40px" Width="48px" BackColor="#C0C0FF">
</asp:DropDownList>Day
<asp:DropDownList ID="Hour1" runat="server" Height="40px" Width="48px" BackColor="#C0C0FF">
</asp:DropDownList>Hour&nbsp;
<asp:DropDownList ID="Min1" runat="server" Height="40px" Width="48px" BackColor="#C0C0FF">
</asp:DropDownList>Minute
                        </P>
<p>
<font size="2"><span style="font-size: 12pt">To:</span> &nbsp;&nbsp; &nbsp;
<asp:DropDownList ID="Year2" runat="server" Height="40px" Width="65px" BackColor="#C0C0FF">
</asp:DropDownList>
Year</font>&nbsp;
<asp:DropDownList ID="Month2" runat="server" Height="40px" Width="48px" BackColor="#C0C0FF">
</asp:DropDownList>Month
<asp:DropDownList ID="Day2" runat="server" Height="40px" Width="48px" BackColor="#C0C0FF">
</asp:DropDownList>Day
<asp:DropDownList ID="Hour2" runat="server" Height="40px" Width="48px" BackColor="#C0C0FF">
</asp:DropDownList>Hour&nbsp;
<asp:DropDownList ID="Min2" runat="server" Height="40px" Width="48px" BackColor="#C0C0FF">
</asp:DropDownList>Minute
</p>
                        <P>
                          <asp:Button id="butRun" Width="56px" runat="server" Text="確定" ForeColor="#0000C0" BorderColor="#00AEEF"
                            BackColor="#E0E0E0" BorderStyle="Solid"></asp:Button>&nbsp;&nbsp;
                          <asp:Button id="butReset" Width="56px" runat="server" Text="修改" ForeColor="#0000C0" BorderColor="#00AEEF"
                            BackColor="#E0E0E0" BorderStyle="Solid"></asp:Button>
                          &nbsp;
<asp:Button ID="butMenu" runat="server" BackColor="#E0E0E0" BorderColor="#00AEEF"
BorderStyle="Solid" ForeColor="#0000C0" Text="回主選單" Width="95px" /><br />
                          <br />
                          <asp:Label id="warning" runat="server" Width="383px" Height="19px" Visible="False" BackColor="#C0FFFF"
                            BorderColor="Transparent" ForeColor="Blue">Start Date error , Please reset.</asp:Label></P>
                      </BLOCKQUOTE>
                    </TD>
                    <CENTER></CENTER>
                  </TR>
                </TABLE>
            <br />
            <br />
          </td>
        </tr>
        <tr>
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
