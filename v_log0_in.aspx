<%@ Page Language="VB" AutoEventWireup="false" CodeFile="v_log0_in.aspx.vb" Inherits="v_log0_in" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    
<script language="javascript" type="text/javascript">
<!--

function IMG1_onclick() {

}

function IMG1_onclick() {

}

function TABLE1_onclick() {

}

// -->
</script>
  <style type="text/css">
    #TABLE1
    {
      width: 669px;
    }
    .style1
    {
      height: 44px;
      width: 334px;
      font-weight: bold;
      text-align: left;
    }
    .style2
    {
      width: 334px;
      text-align: left;
    }
    .style3
    {
      height: 40px;
      width: 334px;
      text-align: left;
    }
    .style4
    {
      height: 31px;
      width: 334px;
      text-align: left;
    }
    .style5
    {
      height: 22px;
      width: 334px;
      text-align: left;
    }
    .style6
    {
      text-align: left;
    }
  </style>
</head>
<body bgcolor="#99ccff" style="text-align: center; position: static; background-color: #ccffff;" topmargin="3">
    <form id="form1" runat="server">
    <div style="background-color: #ccffff">
      <table style="width: 760px">
        <tr>
          <td colspan="3" >
            <asp:Image ID="Image1" runat="server" Height="86px" ImageUrl="~/image/FAR_TOP2.JPG" Width="760px" />
            </td>
        </tr>
        <tr><td colspan="3" style ="text-align: right ">  <asp:Label ID="Label1" runat="server" Text="Label" BackColor="#FFC080" BorderColor="Black" BorderStyle="Solid" BorderWidth="1px" Font-Size="Medium" height="23px"></asp:Label>
        </td></tr>
        <tr>
          <td colspan="3" style="background-color: #ccffff">
             
  &nbsp;<asp:Label ID="lblMDB" runat="server" Font-Size="Medium" ForeColor="Blue" 
              Text="lblMDB" BackColor="#C0FFFF"></asp:Label><strong>&nbsp;</strong><br />
  <strong><span style="font-family: 'Times New Roman'">
      <br />
        主選單</span></strong><br />
<table border="0" cellpadding="0" cellspacing="0" height="95" width="750" bgcolor="#99ccff" align="center">
<tr align="center">
<td style="height: 62px; background-color: #ccffff;">
<table border="0" cellpadding="0" cellspacing="0" height="118" id="TABLE1" 
    language="javascript" onclick="return TABLE1_onclick()">
<tr align="left">
<td class="style1">
  &nbsp;
<asp:Linkbutton ID="hl1" runat="server" Font-Bold="True" Font-Size="Medium" 
    Width="279px">1.電話錄音線路使用統計表</asp:Linkbutton></td>
<td style="width: 111px; height: 44px;" class="style6">
<p align="left">
<span style="font-size: 10pt">&nbsp;&nbsp;
<asp:Linkbutton ID="hl11" runat="server" Font-Bold="True" Font-Size="Medium" 
    Width="258px" Height="22px">11.分行資料維護</asp:Linkbutton></span></p>
</td>
</tr>
<tr align="left">
<td height="40" class="style2">
&nbsp;
<asp:Linkbutton ID="hl2" runat="server" Font-Bold="True" Font-Size="Medium" 
    Width="266px">2.電話錄音進線時段統計表</asp:Linkbutton></td>
<td height="40" style="width: 111px" class="style6">
<b>&nbsp;</b>
    <asp:Linkbutton ID="hl12" runat="server" Font-Bold="True" Font-Size="Medium" 
    Width="236px">12.系統使用者資料維護</asp:Linkbutton></td>
</tr>
    <tr align="left">
        <td height="40" class="style2">
            &nbsp;
<asp:Linkbutton ID="hl3" runat="server" Font-Bold="True" Font-Size="Medium" Width="256px">3.各項服務統計表</asp:Linkbutton></td>
        <td height="40" style="width: 111px" class="style6">
            &nbsp;
    <asp:Linkbutton ID="hl13" runat="server" Font-Bold="True" Font-Size="Medium"  Width="154px">13.群組管理</asp:Linkbutton></td>
    </tr>
    <tr align="left">
        <td height="40" class="style2">
            &nbsp;
    <asp:Linkbutton ID="hl4" runat="server" Font-Bold="True" Width="255px">4.各分行電話錄音統計表</asp:Linkbutton></td>
        <td height="40" style="width: 111px" class="style6">
            &nbsp;
    <asp:Linkbutton ID="hl14" runat="server" Font-Bold="True" Font-Size="Medium" 
              Width="295px">14.使用者&權限異動明細查詢</asp:Linkbutton></td>
    </tr>
<tr align="left">
<td class="style3">
&nbsp;
<asp:Linkbutton ID="hl5" runat="server" Font-Bold="True" Font-Size="Medium" Width="264px">5.各項服務明細查詢</asp:Linkbutton></td>
<td style="width: 111px; height: 40px;" class="style6">
    &nbsp;
<asp:Linkbutton ID="hla" runat="server" Font-Bold="True" Font-Size="Medium" Width="246px">21.變更密碼</asp:Linkbutton></td>
</tr>
<tr>
<td class="style4">
<p class="style6">
  &nbsp;</p>
</td>
<td style="height: 31px; width: 111px;" class="style6">
<p align="left">
&nbsp;&nbsp;</p>
</td>
</tr>
    <tr>
        <td class="style5">
            &nbsp;<asp:Linkbutton ID="hl0" runat="server" Font-Bold="True" Font-Size="Medium" Width="93px" style="table-layout: auto">0.登出　　</asp:Linkbutton></td>
        <td style="width: 111px; height: 22px" class="style6">
        </td>
    </tr>
<tr>
<td colspan="2" style="height: 40px" class="style6">
&nbsp;<br />
 
</td>
</tr>
</table>
<p>
  <font face="新細明體"></font>&nbsp;</p>
  <p>
    <font face="新細明體"></font>&nbsp;</p>
</td>
</tr>
</table>
            <br />
          </td>
        </tr>
        <tr>
          <td style="width: 100px; background-color: #ccffff; height: 21px;">
          </td>
          <td style="width: 100px; background-color: #ccffff; height: 21px;">
          </td>
          <td style="width: 100px; background-color: #ccffff; height: 21px;">
          </td>
        </tr>
      </table>
      <br />
    
    </div>
    </form>
</body>
</html>
