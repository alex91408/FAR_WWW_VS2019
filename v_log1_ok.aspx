<%@ Page Language="VB" AutoEventWireup="false" CodeFile="v_log1_ok.aspx.vb" Inherits="v_log1_ok" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    
<script language="javascript" type="text/javascript">
<!--

function IMG1_onclick() {

}

 //-->
</script>
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
          <td colspan="3" style="height: 285px; background-color: #ccffff;">
<br />
                     <asp:Label ID="lblMDB" runat="server" Font-Size="Large" ForeColor="Blue" Text="lblMDB" BackColor="#C0FFFF"></asp:Label><br />
              <strong>
                  <asp:Label ID="title" runat="server" Text="【1.電話錄音線路使用統計表】"></asp:Label><br />
              </strong>
            <br />
                  <asp:Label id="result1" runat="server" ForeColor="#404040" BackColor="#C0FFFF" Font-Size="Larger">Form:XX To:XX :</asp:Label><br />
            <br />
              &nbsp;&nbsp;
                  <asp:datagrid id="LineGrid" runat="server" ForeColor="DarkOrange" Height="74px" BorderColor="Black"
                    Font-Names="Arial" BackColor="LightSalmon" Font-Bold="False" Font-Italic="False" Font-Overline="False" Font-Strikeout="False" Font-Underline="False" Width="100%">
                    <ItemStyle HorizontalAlign="Right" Height="30px" ForeColor="Black" BorderColor="#FF8000" Width="10%"
                      VerticalAlign="Middle" BackColor="#80FFFF" Font-Bold="False" Font-Italic="False" Font-Overline="False" Font-Strikeout="False" Font-Underline="False"></ItemStyle>
                    <HeaderStyle HorizontalAlign="Center" Height="30px" ForeColor="White" BorderColor="Transparent"
                      Width="10%" VerticalAlign="Middle" BackColor="#00C0C0"></HeaderStyle>
                    <SelectedItemStyle BackColor="#FF8000" Font-Bold="False" Font-Italic="False" Font-Overline="False"
                      Font-Strikeout="False" Font-Underline="False" ForeColor="White" />
                    <AlternatingItemStyle BackColor="#C0FFFF" Font-Bold="False" Font-Italic="False" Font-Overline="False"
                      Font-Strikeout="False" Font-Underline="False" />
                  </asp:datagrid><br />
            &nbsp;<asp:Button id="SaveAsExecl" runat="server" Text="儲存Excel" Width="95px" BackColor="#E0E0E0"
                    ForeColor="#0000C0" BorderColor="#00AEEF" BorderStyle="Solid" Height="25px"></asp:Button>
                &nbsp;
                  <asp:button id="butMenu" runat="server" Width="95px" Text="回主選單" BackColor="#E0E0E0" ForeColor="#0000C0"
                    BorderColor="#00AEEF" BorderStyle="Solid" Height="25px"></asp:button><br />
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
      <br />

   </form>
</body>
</html>
