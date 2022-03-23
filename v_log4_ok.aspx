<%@ Page Language="VB" Debug="true" AutoEventWireup="false" CodeFile="v_log4_ok.aspx.vb" Inherits="v_log4_ok" %>



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    
</head>
<body bgcolor="#99ccff" style="text-align: center; background-color: #ccffff;" topmargin="3">
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
          <td colspan="3" style="height: 21px; background-color: #ccffff;">
<br />
                      <asp:Label ID="lblMDB" runat="server" Font-Size="Large" ForeColor="Blue" Text="lblMDB" BackColor="#C0FFFF"></asp:Label><br />
              <asp:Label ID="title" runat="server" Font-Bold="True" Text="【4.各分行電話錄音統計表】"></asp:Label>
<br />
<br />

                  <asp:Label id="result1" runat="server" ForeColor="#404040" BackColor="#C0FFFF" Font-Size="Larger">FROM XX TO XX
                  </asp:Label><br />
            <br />
              &nbsp; &nbsp;&nbsp;

                  <asp:datagrid id="TimeGrid" runat="server" Height="74px" BorderColor="Black" Font-Names="Arial" BackColor="Black" BorderStyle="Double" BorderWidth="1px" CellPadding="3" GridLines="Horizontal" CellSpacing="1">
                    <ItemStyle HorizontalAlign="Right" Height="30px" ForeColor="#4A3C8C" BorderColor="#0080C0" Width="10%" VerticalAlign="Middle" BackColor="#E7E7FF" Font-Bold="False" Font-Italic="False" Font-Overline="False" Font-Strikeout="False" Font-Underline="False">
                    </ItemStyle>

                    <HeaderStyle HorizontalAlign="Center" Height="30px" ForeColor="#F7F7F7" BorderColor="Transparent"  Width="10%" VerticalAlign="Middle" BackColor="#4A3C8C" Font-Bold="True">
                    </HeaderStyle>
                    <SelectedItemStyle BackColor="#738A9C" Font-Bold="True" ForeColor="#F7F7F7" />
                    <AlternatingItemStyle BackColor="#B5C7DE" />
                      <FooterStyle BackColor="#B5C7DE" ForeColor="#4A3C8C" />
                      <PagerStyle BackColor="#B5C7DE" ForeColor="#4A3C8C" HorizontalAlign="Right" Mode="NumericPages" />
                  </asp:datagrid>
              &nbsp;
                <br />
                  <asp:Button id="SaveAsExecl" runat="server" Text="儲存Excel" Width="95px" BackColor="#E0E0E0" ForeColor="#0000C0" BorderColor="#00AEEF" BorderStyle="Solid" Height="25px"></asp:Button>
                &nbsp;
                  <asp:button id="butMenu" runat="server" Width="95px" Text="回主選單" BackColor="#E0E0E0" ForeColor="#0000C0" BorderColor="#00AEEF" BorderStyle="Solid" Height="25px"></asp:button>
            <br />
          </td>
        </tr>
        <tr style="font-size: 12pt">
          <td style="width: 100px">
          </td>
          <td style="width: 100px">
          </td>
          <td style="width: 100px">
          </td>
        </tr>
      </table>



    </form>


</body>
</html>
