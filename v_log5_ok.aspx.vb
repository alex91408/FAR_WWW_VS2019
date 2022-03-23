'----------------------------------------------------------------------------------------------
' 每次更改請在此區註解    最新放最上面 alex  p400()  
' ver 要改   #define VERSION XXXX   每次異動版本需要修改
'   ver     date    user      meno
'  ------ -------- --------- -------------------------------------------------------

'  V6.3   20220302  vic    "d:\rec_download  -- "e:\rec_download  ,  http://10.48.101.55/rec_download/  --> http://10.48.160.96/rec_download/
'  V6.2   20220222  vic    "e:\rec_download  -- "d:\rec_download  ,  http://10.48.101.55/rec_download/  --> http://10.48.160.96/rec_download/
'  V6.1   20220217  vic    弱掃 20220217

'20220217 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
'Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))



'  V6.0   20220217  alex   弱掃 20220215

''20220217 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc")) ' V6.0   20220215  alex   弱掃 20220215 ; L409 A1 Injection,  Connection String Parameter Pollution
'20220217 new 
'Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))) ' V6.0   20220215  alex   弱掃 20220215 ; L409 A1 Injection,  Connection String Parameter Pollution


Imports System
Imports System.IO

Partial Class v_log5_ok
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load

        Label1.Text = "使用人員：" & Session("acc")

        'Put user code to initialize the page here

        If Not Me.IsPostBack Then

            If Session("Login") = False Then
                Response.Redirect("v_login.aspx")
            End If

            lblMDB.Text = Session("v_menu_1")

            Dim colums As Integer = 15      '共有16個欄位
            Dim c As Integer
            For c = 0 To colums
                GridView1.Columns(c).HeaderStyle.Wrap = False                             '設定Gridview欄位名稱不換行
                GridView1.Columns(c).ItemStyle.Wrap = False                               '設定Gridview欄位列的內容不換行
                GridView1.Columns(c).ItemStyle.HorizontalAlign = HorizontalAlign.Right    '設定Gridview欄位的內容靠右對齊
                GridView1.Columns(c).ItemStyle.Font.Bold = True                           '設定Gridview欄位的字型粗體
            Next

            SaveAsExecl.Visible = False
            '-------------------------------------------------
            ' 970905 ching2 add set0
            Dim set0(150) As String
            Dim iSet As Integer

            set0 = Session("set0")

            Session("v_log5_ok_rec") = 0
            GridView1.Columns(15).Visible = False

            For iSet = 0 To 150
                If set0(iSet) = "" Then
                    Response.Redirect("v_login.aspx")     '避免直接輸入網址跳入
                    'Exit For
                End If

                If set0(iSet) = "v05" Then
                    Exit For
                End If

            Next iSet

            For iSet = 0 To 150
                If set0(iSet) = "v05a" Then
                    SaveAsExecl.Visible = True
                End If
                If set0(iSet) = "v05b" Then
                    GridView1.Columns(15).Visible = True
                End If

            Next iSet

        End If

        '-------------------------------------------------

        '20220223 old SqlDataSource1.ConnectionString = Session("ConnectionString")
        SqlDataSource1.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()))

        ' 改成 --   以月份拆TABLOG  version 3.0 970708 by chris
        '      --   不存在的table表不會加入 Select Command 裡 
        '      --   執行此 Select Command 不會發生資料表不存在的錯誤
        Dim tSQL As String = ""

        '20220217 old Using cn As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc")) ' V6.0   20220215  alex   弱掃 20220215 ; L74 A1 Injection,  Connection String Parameter Pollution
        '20220217 new
        Using cn As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))) ' V6.0   20220215  alex   弱掃 20220215 ; L74 A1 Injection,  Connection String Parameter Pollution
            Dim cmd As New Data.SqlClient.SqlCommand("b_log", cn)
            cmd.CommandType = Data.CommandType.StoredProcedure

            cmd.Parameters.Add(New Data.SqlClient.SqlParameter("@stime", Data.SqlDbType.DateTime))
            cmd.Parameters.Add(New Data.SqlClient.SqlParameter("@etime", Data.SqlDbType.DateTime))
            cmd.Parameters.Add(New Data.SqlClient.SqlParameter("@log_temp", Data.SqlDbType.VarChar))

            cmd.Parameters("@stime").Value = Session("starttime")
            cmd.Parameters("@etime").Value = Session("endtime")
            cmd.Parameters("@log_temp").Value = ""

            cn.Open()
            cmd.CommandTimeout = 500

            Dim er_msg As String = ""
            Try
                tSQL = cmd.ExecuteScalar  '傳回執行預存程序後的return值
                'SqlDataSource1.SelectCommand = ViewState(Session("acc"))
                'GridView1.DataBind()
            Catch ex As Exception
                Dim script As String = ""
                Dim msg As String = ""
                Dim a As Integer
                a = InStr(ex.ToString, "tablog", CompareMethod.Text)
                er_msg = Mid(ex.ToString, a + 7, 6)
                If a > 0 Then
                    msg = Mid(ex.ToString, a, 6)
                End If

                If Left(msg, 6) = "tablog" Then
                    script += "<script>"
                    script += "alert('" & "不能查詢" & er_msg & "月之前的日期" & "');"
                    script += "window.location.href='v_log4_in.aspx';"
                    script += "</script>"
                    '輸出JavaScript() 
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)
                End If

            End Try
            cmd.Dispose()
        End Using

        '-----------------------------------------------------------

        Dim sSQL As String
        REM 取筆數  
        If Len(tSQL) = "0" Then
            Dim script As String = ""

            script += "<script>"
            script += "alert('查詢日期有誤或查詢不到資料 !!');"
            script += "window.location.href='v_log0_in.aspx';"
            script += "</script>"
            '輸出JavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script)
            GridView1.Visible = False
        Else
            sSQL = " Select * from ( "
            sSQL = sSQL & tSQL
            sSQL = sSQL & ") as t1  Where                                                          "
            sSQL = sSQL & "         (t1.stime_in >= '" & Session("starttime") & "' and t1.stime_in <= '" & Session("endtime") & "')"
            sSQL = sSQL & "         and ( t1.trid = t1.trid)       "
            sSQL = sSQL & "         and ( t1.line_hour is Not null  )      "
            sSQL = sSQL & "         and ( t1.stime_in is Not null  )      "
            sSQL = sSQL & "         and ( t1.etime_out is Not null  )      "

            If Session("txtACC_NO") <> "" Then
                sSQL = sSQL & " 	      and ( t1.ao_code Like '%" & Session("txtACC_NO") & "%' )      "
            End If

            If Session("txtPROD_NO") <> "" Then
                sSQL = sSQL & " 	      and ( t1.id_no Like '%" & Session("txtPROD_NO") & "%' )      "
            End If

            If Session("txtSNO") <> "" Then
                sSQL = sSQL & " 	      and ( t1.sno Like '%" & Session("txtSNO") & "%' )      "
            End If


            If Session("txttridno") <> "999" Then
                sSQL = sSQL & " 	      and ( t1.trid Like '%" & Session("txttridno") & "%' )       "
            End If

            If Session("txterrcode") <> "999" Then
                sSQL = sSQL & " 	      and ( t1.error = '" & Session("txterrcode") & "' )      "
            End If

            Select Case Session("SAL05")
                Case 1 To 3
                    If Session("BranchCode") <> "ALL" Then
                        sSQL = sSQL & " 	      and ( t1.bra01 = '" & Session("BranchCode") & "' )      "
                    End If
                Case 4
                    If Session("BranchCode") <> "ALL" Then
                        sSQL = sSQL & " 	      and ( t1.bra01 = '" & Session("BranchCode") & "' )      "
                    Else
                        sSQL = sSQL & " 	      and ( t1.bra05 = '" & Session("BRA05") & "' )      "
                    End If
            End Select




            sSQL = sSQL & "                                                                "
            sSQL = sSQL & " Order by t1.stime_in desc"

            '定義Select指令
            SqlDataSource1.SelectCommand = sSQL

            result1.Text = "From:" & Session("starttime") & "  -  " & Session("endtime")

            Dim ii As New DataSourceSelectArguments
            Dim v As Data.DataView = CType(SqlDataSource1.Select(ii), Data.DataView)
            Lbl2.Text = "共:" & v.Count & " 筆 "

        End If

    End Sub

    Private Sub butMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butMenu.Click
        '設定 session 表示使用者已經 login
        Session("Login") = True
        '轉到主選單
        '- 呼叫預存程序 
        Dim tSQL As String = ""
        '20220217 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc")) ' V6.0   20220215  alex   弱掃 20220215 ; L200 A1 Injection,  Connection String Parameter Pollution
        '20220217 new 
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))) ' V6.0   20220215  alex   弱掃 20220215 ; L200 A1 Injection,  Connection String Parameter Pollution
            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "5.各項服務明細查詢->回主選單"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using


        Response.Redirect("v_log0_in.aspx")
    End Sub

    Private Sub SaveAsExecl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsExecl.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""
        '20220217 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc")) ' V6.0   20220215  alex   弱掃 20220215 ; L225 A1 Injection,  Connection String Parameter Pollution
        '20220217 new 
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))) ' V6.0   20220215  alex   弱掃 20220215 ; L225 A1 Injection,  Connection String Parameter Pollution
            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "5.各項服務明細查詢-save excel"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        '20220223 old SqlDataSource1.ConnectionString = Session("ConnectionString")
        SqlDataSource1.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()))

        Dim style As String = "<style> .text { mso-number-format:\@; } </script> "
        Dim sw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(sw)


        Response.Clear()        
        Response.Charset = "big5"  ' 在2003EXCEL和2000EXCEL, 中文都不會變亂碼
        Response.ContentType = "Content-Language;content=zh-tw"   ' 新加的
        Response.ContentType = "application/vnd.ms-excel"
        Response.AppendHeader("Content-Disposition", "attachment; filename=Report_05.xls")

        Response.Write("<meta http-equiv=Content-Type content=text/html")
        Response.ContentEncoding = System.Text.UTF8Encoding.UTF8

        GridView1.Columns(15).Visible = False
        GridView1.AllowPaging = False
        GridView1.AllowSorting = False

        GridView1.DataBind()

        Response.Buffer = True

        title.RenderControl(hw)
        result1.RenderControl(hw)
        GridView1.RenderControl(hw)
        '980525 ching2 add erro
        lblErr1.RenderControl(hw)
        lblErr2.RenderControl(hw)
        lblErr3.RenderControl(hw)

        Response.Write(style)
        Response.Write(sw.ToString())
        Response.End()
    End Sub

    'Protected Sub GridView1_RowCommand(ByVal source As Object, ByVal e As GridViewCommandEventArgs)
    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand


        If e.CommandName = "DownloadRec" Then      'listen

            Dim index As Integer = Convert.ToInt32(e.CommandArgument)
            Dim strIdx As String
            Dim row As GridViewRow = GridView1.Rows(index)

            strIdx = row.Cells(12).Text

            '- 呼叫預存程序 
            Dim tSQL As String = ""
            '20220217 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc")) ' V6.0   20220215  alex   弱掃 20220215 ; L293 A1 Injection,  Connection String Parameter Pollution
            '20220217 new 
            Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))) ' V6.0   20220215  alex   弱掃 20220215 ; L293 A1 Injection,  Connection String Parameter Pollution
                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "5.各項服務明細查詢-下載" + strIdx

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

            'Dim test As String = Page.Request.UserHostAddress.ToString()

            'Response.Write("<script>window.open('" + Server.UrlEncode(strIdx) + "')</script>")

            '980525 ching2 test ok
            'Response.Write("<script>window.open(' http://10.50.1.201" + strIdx + "')</script>")

            'FileCopy("d:\rec\" & Trim(row.Cells(8).Text) & "\" & row.Cells(12).Text.Substring(11), "d:\rec\" & Trim(row.Cells(8).Text) & ".mp3")
            'Response.Write("<script>window.open(' http://192.168.1.177/rec/" & Trim(row.Cells(8).Text) & ".mp3" & "')</script>")
            ''Kill("d:\rec\" & Trim(row.Cells(8).Text) & ".mp3")


            'If Dir("e:\rec_download\" & Session("acc") & "_rec.mp3") = Session("acc") & "_rec.mp3" Then
            Dim file1 as String

            '20220222 old  file1 = Dir("e:\rec_download\" & Session("acc") & "_*_rec.mp3")
            '  V6.3   20220302  vic    "d:\rec_download  -- "e:\rec_download  ,  http://10.48.101.55/rec_download/  --> http://10.48.160.96/rec_download/
            '20220302 old file1 = Dir("d:\rec_download\" & Session("acc") & "_*_rec.mp3")
            file1 = Dir("e:\rec_download\" & Session("acc") & "_*_rec.mp3")
            If len(file1) > 0 Then
                '1001012 old
                '   Kill("e:\rec_download\" & Session("acc") & "_rec.mp3")
                '20220222 old  Kill("e:\rec_download\" & Session("acc") & "_*_rec.mp3")
                '  V6.3   20220302  vic    "d:\rec_download  -- "e:\rec_download  ,  http://10.48.101.55/rec_download/  --> http://10.48.160.96/rec_download/
                '20220302 old Kill("d:\rec_download\" & Session("acc") & "_*_rec.mp3")
                Kill("e:\rec_download\" & Session("acc") & "_*_rec.mp3")
            End If

            'If Dir("D:\rec_download\" & Session("acc") & "_rec.mp3") = Session("acc") & "_rec.mp3" Then
            '    Kill("D:\rec_download\" & Session("acc") & "_rec.mp3")
            'End If


            If row.Cells(12).Text <> "-" Then

            '1001012 old
            '    FileCopy("e:\rec\" & Trim(row.Cells(8).Text) & "\" & row.Cells(13).Text.Substring(11), "e:\rec_download\" & Session("acc") & "_rec.mp3")
            '    Response.Write("<script>window.open(' http://10.50.1.201/rec_download/" & Session("acc") & "_rec.mp3" & "')</script>")
            
            '1001012 new
            Dim a as string
            a = hour(now) & minute(now) & second(now)

                '20220222 old  FileCopy("e:\rec\" & Trim(row.Cells(8).Text) & "\" & row.Cells(13).Text.Substring(11), "e:\rec_download\" & Session("acc") & "_" & a & "_rec.mp3")
                '  V6.3   20220302  vic    "d:\rec_download  -- "e:\rec_download  ,  http://10.48.101.55/rec_download/  --> http://10.48.160.96/rec_download/
                '20220302 old FileCopy("d:\rec\" & Trim(row.Cells(8).Text) & "\" & row.Cells(13).Text.Substring(11), "d:\rec_download\" & Session("acc") & "_" & a & "_rec.mp3")
                FileCopy("e:\rec\" & Trim(row.Cells(8).Text) & "\" & row.Cells(13).Text.Substring(11), "e:\rec_download\" & Session("acc") & "_" & a & "_rec.mp3")




                '20220222 old  Response.Write("<script>window.open(' http://10.48.101.55/rec_download/" & Session("acc") & "_" & a & "_rec.mp3" & "')</script>")
                Response.Write("<script>window.open(' http://10.48.160.96/rec_download/" & Session("acc") & "_" & a & "_rec.mp3" & "')</script>")


                'a$ = "<script>window.open('http://10.48.101.55/rec_download/" & Session("acc") & "_" & a & "_rec.mp3" & "')</script>"

                'Response.Write(a$)
                'Response.Redirect(a$)


                'Response.Write("<script>window.open(' http://127.0.0.1/rec_download/" & Session("acc") & "_rec.mp3" & "')</script>")
                'FileCopy("D:\rec\" & Trim(row.Cells(8).Text) & "\" & row.Cells(13).Text.Substring(11), "D:\rec_download\" & Session("acc") & "_rec.mp3")
                'Response.Write("<script>window.open(' http://127.0.0.1/rec_download/" & Session("acc") & "_rec.mp3" & "')</script>")

            End If


        End If

    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        If e.Row.RowIndex <> -1 Then

            'CType(e.Row.Cells(0).Controls(0), Button).Attributes.Add("OnClick", "return confirm('Are you sure resend fax?');") '= strJava

            'CType(e.Row.Cells(0).Controls(0), Button).Attributes.Add("OnClick", "if ( !confirm('Are you sure resend fax?')) {return false};")
            'CType(e.Row.Cells(1).Controls(0), Button).Attributes.Add("OnClick", "if ( !confirm('Are you sure resend mail?')) {return false};")
        End If

        If e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.Cells(0).Attributes.Add("class", "text")
            e.Row.Cells(1).Attributes.Add("class", "text")
            e.Row.Cells(2).Attributes.Add("class", "text")
            e.Row.Cells(3).Attributes.Add("class", "text")
            e.Row.Cells(4).Attributes.Add("class", "text")
            e.Row.Cells(5).Attributes.Add("class", "text")
            e.Row.Cells(6).Attributes.Add("class", "text")
            e.Row.Cells(7).Attributes.Add("class", "text")
            e.Row.Cells(8).Attributes.Add("class", "text")
            e.Row.Cells(9).Attributes.Add("class", "text")
            e.Row.Cells(10).Attributes.Add("class", "text")
            e.Row.Cells(11).Attributes.Add("class", "text")
            e.Row.Cells(12).Attributes.Add("class", "text")
            e.Row.Cells(13).Attributes.Add("class", "text")
            e.Row.Cells(14).Attributes.Add("class", "text")
            e.Row.Cells(15).Attributes.Add("class", "text")

        End If
        If e.Row.RowIndex = GridView1.EditIndex And e.Row.RowIndex <> -1 Then
            CType(e.Row.Cells(33).Controls(0), TextBox).Width = Unit.Point(100) 'Faxtel
            CType(e.Row.Cells(34).Controls(0), TextBox).Width = Unit.Point(20) 'Period
            CType(e.Row.Cells(35).Controls(0), TextBox).Width = Unit.Point(20) 'Fax_times
            CType(e.Row.Cells(36).Controls(0), TextBox).Width = Unit.Point(20) 'maxtimes
            CType(e.Row.Cells(37).Controls(0), TextBox).Width = Unit.Point(50) 'Faxerrcode
            CType(e.Row.Cells(38).Controls(0), TextBox).Width = Unit.Point(20) 'EmailFlag
        End If
    End Sub

    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        '處理'GridView' 的控制項 'GridView' 必須置於有 runat=server 的表單標記之中
        'Dim a = "1"
    End Sub
    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""
        '20220217 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc")) ' V6.0   20220215  alex   弱掃 20220215 ; L409 A1 Injection,  Connection String Parameter Pollution
        '20220217 new 
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))) ' V6.0   20220215  alex   弱掃 20220215 ; L409 A1 Injection,  Connection String Parameter Pollution
            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "0.登出"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Session("acc") = ""

        Response.Redirect("V_Login.aspx")
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

    End Sub

   
End Class


