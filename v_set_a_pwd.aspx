<%@ Page Language="VB" AutoEventWireup="false" CodeFile="v_set_a_pwd.aspx.vb" Inherits="v_set5_pwd" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    
  <style type="text/css">
    .style1
    {
      font-family: �ө���;
    }
    .style2
    {
      line-height: 150%;
      font-family: �ө���;
      margin: 0px 4px;
    }
    .style3
    {
      font-size: 12pt;
      font-family: �ө���;
    }
  </style>
    
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
      <table style="width: 760px">
        <tr>
          <td colspan="3">
            <asp:Image ID="Image1" runat="server" Height="86px" ImageUrl="~/image/FAR_TOP2.JPG"
              Width="760px" /></td>
        </tr>
        <tr><td colspan="3" style ="text-align: right ">  
            <asp:Label ID="Label1" runat="server" Text="Label" BackColor="#FFC080" BorderColor="Black" BorderStyle="Solid" BorderWidth="1px" Font-Size="Medium" height="23px"></asp:Label>
            <asp:button id="Button1" runat="server" Width="50px" height="25px" Text="�n�X" ForeColor="Black" BorderColor="black" Font-Size="Medium"
                BackColor="#E0E0E0" BorderStyle="Solid" BorderWidth="1px"></asp:button>
        </td></tr>
        <tr>
          <td colspan="3" style="background-color: #ccffff">
            <br />
                  <asp:Label ID="lblMDB" runat="server" Font-Size="Large" ForeColor="Blue" Text="lblMDB" Visible = "False" ></asp:Label><br />
              �i<strong>21.�ܧ�K�X</strong>�j<br />
            <br />
              <table style="WIDTH: 729px; HEIGHT: 108px; background-color: black; margin-right: 29px;" 
              height="108" cellSpacing="1" cellPadding="1" bgColor="#0080c0" border="0">
                <tr bgColor="#ffffff">
                  <td bgColor="#0080c0" height="22" 
                    style="width: 550px; background-color: #ccffff;" class="style1">
                    <p class="style2" align="center">&nbsp;</p>
                  </td>
                </tr>

                <tr bgColor="#ffffff" style="font-family: Times New Roman">
                  <td vAlign="middle" align="left" 
                    style="height: 204px; width: 550px; background-color: #ccffff;" 
                    bgcolor="#99ccff" class="style1">

                    <blockquote style="width: 596px; text-align: center">
                      <br />
                      <FONT face="�s�ө���">
                      <span style="font-size: 12pt; mso-bidi-font-family: 'Times New Roman';
                            mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                            mso-bidi-language: AR-SA" class="style1">�ϥΪ̱b��</span></FONT><span 
                        class="style1">�G </span><FONT face="�s�ө���">
                        <asp:textbox id="txtAcc" Width="120px" runat="server" Height="22px" 
                        Enabled="False" BackColor="#C0C0FF" CssClass="style1"></asp:textbox>
                      <span class="style1">
<br />
<br />
                        </span>
                        <span style="font-size: 12pt; mso-bidi-font-family: 'Times New Roman';
                            mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                            mso-bidi-language: AR-SA" class="style1">�±K�X</span><span class="style1">�G &nbsp;&nbsp; &nbsp;</span><asp:textbox 
                        id="txtPwd" Width="120px" runat="server" Height="22px" TextMode="Password" 
                        BackColor="#C0C0FF" MaxLength="8" CssClass="style1"></asp:textbox>
                      <span class="style1"><br />
<br />
                        </span>
                        <span style="font-size: 12pt; mso-bidi-font-family: 'Times New Roman';
                            mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                            mso-bidi-language: AR-SA" class="style1">�s�K�X</span><span class="style1">�G &nbsp;&nbsp; &nbsp;</span><asp:textbox 
                        id="txtPwd2" Width="120px" runat="server" Height="22px" TextMode="Password" 
                        BackColor="#C0C0FF" MaxLength="8" CssClass="style1"></asp:textbox>
                      <br />
                      <br />
                      <span><span class="style1">�T�{�s�K�X�G </span>&nbsp;<asp:TextBox ID="txtPwd3" 
                        runat="server" Height="22px" TextMode="Password" Width="120px" 
                        BackColor="#C0C0FF" MaxLength="8" CssClass="style1"></asp:TextBox>&nbsp;</span></FONT></blockquote>
<p style="width: 636px; text-align: center">
  <span><span class="style1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span>
  <span class="style1"><FONT face="�s�ө���">
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </FONT>
    <br />
    
&nbsp; &nbsp; &nbsp;&nbsp;
                        
                  </span>
                        
                  <asp:button id="cmdRun" Width="56px" runat="server" Text="�T�w" 
    ForeColor="#0000C0" BorderColor="#00AEEF"
                          BorderStyle="Solid" BackColor="#E0E0E0" 
    CssClass="style1"></asp:button><span class="style1">&nbsp; </span>
                        
                        <asp:button id="cmdClear" Width="112px" runat="server" 
    Text="���s�]�w" ForeColor="#0000C0" BorderColor="#00AEEF"
                          BorderStyle="Solid" BackColor="#E0E0E0" 
    CssClass="style1"></asp:button><span class="style1">&nbsp; </span>
                  <asp:Button id="butMenu" Width="95px" runat="server" 
    Text="�^�D���" ForeColor="#0000C0" BorderColor="#00AEEF"
                    BorderStyle="Solid" BackColor="#E0E0E0" CssClass="style1"></asp:Button></p>
                      <p class="style3" 
                      style="mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA">
                          &nbsp; &nbsp;&nbsp; ���G�K�X�ܧ�W�h</p>
                      <p>
                          <span lang="EN-US" style="font-size: 12pt; mso-bidi-font-family: 'Times New Roman';
                              mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                              mso-bidi-language: AR-SA" class="style1">&nbsp; &nbsp; &nbsp;1.</span><span style="font-size: 12pt; font-family: �ө���;
                                  mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US;
                                  mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA">���i�ɾ��έ����A678</span>��<span
                                      lang="EN-US">87654321</span></span></p>
                      <p>
                          <span lang="EN-US" 
                            style="font-size: 12pt; font-family: �ө���; mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA">&nbsp; &nbsp; &nbsp;2.<span style="font-size: 12pt;
                                  font-family: �ө���; mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt;
                                  mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA">���i�s��A<span
                                      lang="EN-US">11111111</span>��<span lang="EN-US">22222222</span></span></span></p>
                      <p>
                          <span class="style1">&nbsp; &nbsp; &nbsp;3.</span><span style="font-size: 12pt; mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US;
                              mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA" class="style1">���i�M�±K�X�ۦP</span></p>
                      <p>
                          <span class="style1">&nbsp; &nbsp; &nbsp;4.</span><span style="font-size: 12pt; font-family: �ө���;
                              mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US;
                              mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA">�K�X���׬�<span lang="EN-US">8</span>��A��<span
                                  lang="EN-US">0~9</span>�Ʀr�զX�Ӧ�</span></p>
                      <p>
  <br class="style1" />
                        <span class="style1">&nbsp; &nbsp; &nbsp;<asp:Label id="lblStat" runat="server" Height="22px" Width="426px" 
                          Visible="False" ForeColor="Red">lblStat</asp:Label><BR>
                        </span>
                      </p>
                  </td>
                  
                </tr>
              </table>
                  &nbsp;<br />
            <br />
            <br />
          </td>
        </tr>
        <tr>
          <td colspan="3" style="height: 21px">
          </td>
        </tr>
      </table>
      <br />
      <br />
      <br />
      <br />
      <br />
      <br />
    
  
    </form>
</body>
</html>
