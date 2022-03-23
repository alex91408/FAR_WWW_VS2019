<%@ Page Language="VB" AutoEventWireup="false" CodeFile="v_log14_in.aspx.vb" Inherits="v_log14_in" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>VCRO IVR SYSTEM</title>
</head>

  <body style="text-align: center; font-size: 12pt; font-family: Times New Roman; background-color: #ccffff;" bgcolor="#99ccff" leftmargin="3" topmargin="3">


    <form id="Form2" method="post" runat="server">
      <table style="width: 760px">
        <tr>
          <td colspan="3" style="height: 21px">
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
              &nbsp;<br/>
                  <asp:Label ID="lblMDB" runat="server" Font-Size="Large" ForeColor="Blue" Text="lblMDB" BackColor="#C0FFFF"></asp:Label><br />
              <strong>【14.<span style="font-size: 12pt; font-family: 細明體; mso-bidi-font-family: 'Times New Roman';
                  mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                  mso-bidi-language: AR-SA">使用者<span lang="EN-US">&amp;</span>權限異動明細查詢</span>】</strong><br />
            <br />
            <table id="Table2" bgcolor="#0080c0" border="0" cellpadding="1" cellspacing="1" height="222"
              style="height: 222px; background-color: black;" width="600">
              <tr bgcolor="#ffffff">
                <td bgcolor="#0080c0" style="width: 567px; height: 22px; background-color: #ccffff;">
                  <p align="center" class="style5" style="margin: 0px 4px; line-height: 150%">
                    <font face="新細明體" size="2"><b>&nbsp;<span style="font-size: 12pt; font-family: Times New Roman">
                        輸入查詢時間</span></b></font></p>
                </td>
              </tr>
              <tr bgcolor="#ffffff" style="font-size: 12pt; font-family: Times New Roman">
                <td align="left" bgcolor="#99ccff" style="width: 567px; height: 49px; background-color: #ccffff;" valign="middle">
                  <blockquote>
                    <p>
                      <span style="font-size: 10pt">
                        <br />
                          From:&nbsp;</span>&nbsp;<asp:DropDownList ID="Year1" runat="server" BackColor="#C0C0FF" Height="40px" Width="65px">
                      </asp:DropDownList>
                        <span style="font-size: 10pt">Year&nbsp; </span>
                      <asp:DropDownList ID="Month1" runat="server" BackColor="#C0C0FF" Height="40px" Width="48px">
                      </asp:DropDownList><span style="font-size: 10pt">Month</span>
                      <asp:DropDownList ID="Day1" runat="server" BackColor="#C0C0FF" Height="40px" Width="48px">
                      </asp:DropDownList><span style="font-size: 10pt">Day</span>
                      <asp:DropDownList ID="Hour1" runat="server" BackColor="#C0C0FF" Height="40px" Width="48px">
                      </asp:DropDownList><span style="font-size: 10pt">Hour</span>&nbsp;
                      <asp:DropDownList ID="Min1" runat="server" BackColor="#C0C0FF" Height="40px" Width="48px">
                      </asp:DropDownList><span style="font-size: 10pt">Minute</span>
                    </p>
                    <p>
                      <font size="2"><span style="font-size: 12pt"><span style="font-size: 10pt">To</span>:</span> &nbsp;&nbsp; &nbsp;
                        <asp:DropDownList ID="Year2" runat="server" BackColor="#C0C0FF" Height="40px" Width="65px">
                        </asp:DropDownList>
                          <span style="font-size: 10pt">Year</span></font>&nbsp;
                      <asp:DropDownList ID="Month2" runat="server" BackColor="#C0C0FF" Height="40px" Width="48px">
                      </asp:DropDownList><span style="font-size: 10pt">Month</span>
                      <asp:DropDownList ID="Day2" runat="server" BackColor="#C0C0FF" Height="40px" Width="48px">
                      </asp:DropDownList><span style="font-size: 10pt">Day</span>
                      <asp:DropDownList ID="Hour2" runat="server" BackColor="#C0C0FF" Height="40px" Width="48px">
                      </asp:DropDownList><span style="font-size: 10pt">Hour</span>&nbsp;
                      <asp:DropDownList ID="Min2" runat="server" BackColor="#C0C0FF" Height="40px" Width="48px">
                      </asp:DropDownList><span style="font-size: 10pt">Minute</span>
                    </p>
                    
                    <p>
                        <span style="font-family: 細明體">分行代碼 &nbsp; &nbsp;
                          &nbsp;&nbsp; &nbsp;<asp:DropDownList ID="BranchCode" runat="server" BackColor="#C0C0FF"
                              Height="40px" Width="252px">
                          </asp:DropDownList></span> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
                    </p>
                      <p>
                          <span>
                        使用者帳號&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                          <asp:textbox id="txtACC_NO" runat="server" Width="244px" BackColor="#C0C0FF" 
                            MaxLength="6"></asp:textbox></span></p>
                      <p>
                      <asp:Button ID="butRun" runat="server" BackColor="#E0E0E0" BorderColor="#00AEEF"
                        BorderStyle="Solid" ForeColor="#0000C0" Text="確定" Width="56px" />
                      &nbsp;&nbsp;<asp:Button ID="butReset" runat="server" BackColor="#E0E0E0" BorderColor="#00AEEF"
                        BorderStyle="Solid" ForeColor="#0000C0" Text="修改" Width="56px" />
                      &nbsp;&nbsp;<asp:Button ID="butMenu" runat="server" BackColor="#E0E0E0" BorderColor="#00AEEF"
                        BorderStyle="Solid" ForeColor="#0000C0" Text="回主選單" Width="95px" /></p>
                      <p>
                          &nbsp;<asp:Label ID="warning" runat="server" BackColor="#C0FFFF" BorderColor="Transparent"
                        ForeColor="Blue" Height="19px" Visible="False" Width="383px">Start Date error , Please reset.</asp:Label> &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</p>
                   
                  </blockquote>
                </td>
               
              </tr>
               
            </table>
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



    </form>

  </body>


</html>
