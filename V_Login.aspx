<%@ Page Language="VB" AutoEventWireup="false" CodeFile="V_Login.aspx.vb" Inherits="V_Login" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    
</head>

<script type="text/javascript">

function choiceUser(fieldClientID,isMutilSelect){
window.open('UserSelector.aspx?fieldClientID='+fieldClientID.id+'&mutilSelect='+isMutilSelect,'','width=220,height=220','');
}
function choiceDate(fieldClientID){
window.open('dtPicker.aspx?fieldClientID='+fieldClientID.id,'','width=220,height=220','');
}

</script>



<body bgcolor="#99ccff" style="text-align: center; background-color: #ccffff;" topmargin="3">
  <form id="form1" runat="server">
    <table style="width: 800px">
      <tr>
        <td colspan="4" style="text-align: center" align="left">
            <img src="image/FAR_TOP2.JPG" Width="760px" Height="86px"/>&nbsp;</td>
      </tr>
    </table>
      <br />
      <table style="width: 400px">
          <tr>
              <td style="width: 2261px; height: 38px" align="right">
                  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                  &nbsp; &nbsp; &nbsp;<img src="image/FAR_logo1.JPG" style="width: 369px; height: 194px" /></td>
              <td style="width: 164px; height: 38px">
              </td>
              <td style="width: 407px; height: 38px; text-align: left;" align="left">
                  <strong><span style="font-size: 16pt; font-family: 標楷體">
                      <br />
                      <br />
                遠東商銀電話錄音系統<br />
                  </span></strong>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
            &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<br />
            &nbsp;<img height="12" src="image/arrow6.gif" width="27" />&nbsp;<font size="3"><strong>
  <asp:Label ID="lblMDB" runat="server" Font-Size="Medium" ForeColor="Blue" Text="lblMDB" Visible="False"></asp:Label><br />
</strong></font>
            <br />
                  <span style="font-family: 標楷體">帳號: &nbsp; </span>
                  <asp:TextBox ID="txtAcc" runat="server" Style="z-index: 106; left: 182px;
top: 192px" Width="120px" BackColor="#C0C0FF" BorderColor="White" MaxLength="6"></asp:TextBox>
        &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;<br />
          <br />
                  <span style="font-family: 標楷體">密碼: &nbsp; </span>
                  <asp:TextBox ID="txtPwd" runat="server" Height="18px" Style="z-index: 107; left: 184px; top: 250px" Width="120px" BackColor="#C0C0FF" TextMode="Password" MaxLength="8">1234</asp:TextBox>
        &nbsp; &nbsp; 
        
        <asp:Button ID="butLogin" runat="server" BorderColor="#00AEEF" BorderStyle="Solid"
ForeColor="#0000C0" Style="z-index: 108; left: 324px;
top: 252px" Text="登入" Width="72px" Font-Size="Medium" />
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;&nbsp;&nbsp; &nbsp; &nbsp;
          &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<br />
          <br />
          &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;<asp:Label ID="lblStat" runat="server" ForeColor="Blue" Height="22px" Text="lblStat"
Width="487px" Font-Size="Medium"></asp:Label><br />
              </td>
          </tr>
          <tr>
              <td style="width: 2261px" align="right">
              </td>
              <td style="width: 164px">
              </td>
              <td style="width: 507px" align="left">
              </td>
          </tr>
          <tr>
              <td style="width: 2261px" align="right">
              </td>
              <td style="width: 164px">
              </td>
              <td style="width: 507px" align="left">
              </td>
          </tr>
      </table>
    <br />
    </form>
    
</body>
</html>
