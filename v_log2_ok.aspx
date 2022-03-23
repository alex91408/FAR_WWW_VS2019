<%@ Page Language="VB" AutoEventWireup="false" CodeFile="v_log2_ok.aspx.vb" Inherits="v_log2_ok" %>

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
          <td colspan="3" style="height: 339px">
            <br />
                       <asp:Label ID="lblMDB" runat="server" Font-Size="Large" ForeColor="Blue" Text="lblMDB" BackColor="#C0FFFF"></asp:Label><br />
              <strong><span style="font-size: 12pt; font-family: 細明體; mso-bidi-font-family: 'Times New Roman';
                  mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                  mso-bidi-language: AR-SA">
                  <asp:Label ID="title" runat="server" Font-Bold="True" Text="【2.電話錄音進線時段統計表】"></asp:Label></span></strong><P><asp:label id="result1" runat="server" Font-Size="Larger" BackColor="#C0FFFF" ForeColor="#404040">Form XX To XX  </asp:label>
                  <br />
                  <br />
                  <asp:datagrid id="TimeGrid" runat="server" ForeColor="Black" Font-Names="Arial" BorderColor="#FF8000" Height="74px" BackColor="#FF8000">
                    <ItemStyle HorizontalAlign="Right" Height="30px" ForeColor="Black" BorderColor="#0080C0" Width="10%" VerticalAlign="Middle" BackColor="#80FFFF" Font-Bold="False" Font-Italic="False" Font-Overline="False" Font-Strikeout="False" Font-Underline="False"></ItemStyle>
                    <HeaderStyle HorizontalAlign="Center" Height="30px" ForeColor="White" BorderColor="Transparent" Width="10%" VerticalAlign="Middle" BackColor="#00C0C0"></HeaderStyle>
                    <EditItemStyle BackColor="#FF8000" Font-Bold="False" Font-Italic="False" Font-Overline="False"
                      Font-Strikeout="False" Font-Underline="False" />
                    <SelectedItemStyle BackColor="#99CCFF" Font-Bold="False" Font-Italic="False" Font-Overline="False"
                      Font-Strikeout="False" Font-Underline="False" />
                    <AlternatingItemStyle BackColor="#C0FFFF" Font-Bold="False" Font-Italic="False" Font-Overline="False"
                      Font-Strikeout="False" Font-Underline="False" />
                  </asp:datagrid>
                </P>

                <P><asp:button id="SaveAsExecl" runat="server" Width="95px" Text="儲存Excel" ForeColor="#0000C0" BackColor="#E0E0E0" BorderColor="#00AEEF" BorderStyle="Solid" Height="25px">
                    </asp:button>&nbsp;&nbsp;
                  <asp:button id="butMenu" runat="server" Width="95px" Text="回主選單" ForeColor="#0000C0" BackColor="#E0E0E0"
                    BorderColor="#00AEEF" BorderStyle="Solid" Height="25px"></asp:button>
                </P>
            <p>
              &nbsp;</p>
          </td>
        </tr>
        <tr style="font-size: 12pt">
          <td style="width: 100px; height: 29px;">
          </td>
          <td style="width: 100px; height: 29px;">
          </td>
          <td style="width: 100px; height: 29px;">
          </td>
        </tr>
      </table>
      <br />


    </form>
</body>
</html>
