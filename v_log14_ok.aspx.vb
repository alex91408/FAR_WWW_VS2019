'----------------------------------------------------------------------------------------------
' 每次更改請在此區註解    最新放最上面 alex  p400()  
' ver 要改   #define VERSION XXXX   每次異動版本需要修改
'   ver     date    user      meno
'  ------ -------- --------- -------------------------------------------------------

'  V6.0   20220217  alex   弱掃 20220215

'20220217 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc")) '  V6.0   20220215  alex   弱掃 20220215 ; L97 A1 Injection,  Connection String Parameter Pollution
'20220217 new
'Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))) '  V6.0   20220215  alex   弱掃 20220215 ; L97 A1 Injection,  Connection String Parameter Pollution
'20220217 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc")) '  V6.0   20220215  alex   弱掃 20220215 ; L121 A1 Injection,  Connection String Parameter Pollution
'20220217 new 
'Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))) '  V6.0   20220215  alex   弱掃 20220215 ; L121 A1 Injection,  Connection String Parameter Pollution
'20220217 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc")) '  V6.0   20220215  alex   弱掃 20220215 ; L208 A1 Injection,  Connection String Parameter Pollution
'20220217 new 
'Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))) '  V6.0   20220215  alex   弱掃 20220215 ; L208 A1 Injection,  Connection String Parameter Pollution


Imports System
Imports System.IO

Partial Class v_log14_ok
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load

        Label1.Text = "使用人員：" & Session("acc")

        'Put user code to initialize the page here

        'If Not Me.IsPostBack Then

        If Session("Login") = False Then
            Response.Redirect("v_login.aspx")
        End If

        SaveAsExecl.Visible = False
        '-------------------------------------------------
        ' 970905 ching2 add set0
        Dim set0(150) As String
        Dim iSet As Integer

        set0 = Session("set0")

        For iSet = 0 To 150
            If set0(iSet) = "" Then
                Response.Redirect("v_login.aspx")     '避免直接輸入網址跳入
                'Exit For
            End If

            If set0(iSet) = "v14" Then
                Exit For
            End If

        Next iSet

        For iSet = 0 To 150
            If set0(iSet) = "v14a" Then
                SaveAsExecl.Visible = True
            End If

        Next iSet
        '-------------------------------------------------

        lblMDB.Text = Session("v_menu_1")
        Session("SaveAsExecl_flag") = "0"

        '20220223 old SqlDataSource1.ConnectionString = Session("ConnectionString")
        SqlDataSource1.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()))

        Dim sSQL As String
        REM 取筆數  

        sSQL = " Select * from LogTrace Where "
        sSQL = sSQL & " (ActionTime >= '" & Session("starttime") & "' and ActionTime <= '" & Session("endtime") & "')"

        If Session("txtACC_NO") <> "" Then
            sSQL = sSQL & " 	      and ( AccName Like '%" & Session("txtACC_NO") & "%' )      "
        End If

        Select Case Session("SAL05")
            Case 1 To 3
                If Session("BranchCode") <> "ALL" Then
                    sSQL = sSQL & " 	      and ( bra01 = '" & Session("BranchCode") & "' )      "
                End If
            Case 4
                If Session("BranchCode") <> "ALL" Then
                    sSQL = sSQL & " 	      and ( bra01 = '" & Session("BranchCode") & "' )      "
                Else
                    sSQL = sSQL & " 	      and ( bra05 = '" & Session("BRA05") & "' )      "
                End If
        End Select


        sSQL = sSQL & "                                                                "
        sSQL = sSQL & " Order by ActionTime desc"

        '定義Select指令
        SqlDataSource1.SelectCommand = sSQL

        result1.Text = "From:" & Session("starttime") & "  -  " & Session("endtime")

        Dim ii As New DataSourceSelectArguments
        Dim v As Data.DataView = CType(SqlDataSource1.Select(ii), Data.DataView)
        Lbl2.Text = "共:" & v.Count & " 筆 "

        ' End If

    End Sub

    Private Sub butMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butMenu.Click
        '設定 session 表示使用者已經 login
        Session("Login") = True
        '轉到主選單
        '- 呼叫預存程序 
        Dim tSQL As String = ""
        '20220217 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc")) '  V6.0   20220215  alex   弱掃 20220215 ; L97 A1 Injection,  Connection String Parameter Pollution
        '20220217 new
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))) '  V6.0   20220215  alex   弱掃 20220215 ; L97 A1 Injection,  Connection String Parameter Pollution

            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "14.使用者&權限異動明細查詢->回主選單"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_log0_in.aspx")
    End Sub

    Private Sub SaveAsExecl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsExecl.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""
        '20220217 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc")) '  V6.0   20220215  alex   弱掃 20220215 ; L121 A1 Injection,  Connection String Parameter Pollution
        '20220217 new 
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))) '  V6.0   20220215  alex   弱掃 20220215 ; L121 A1 Injection,  Connection String Parameter Pollution
            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "14.使用者&權限異動明細查詢-save execl"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Dim style As String = "<style> .text { mso-number-format:\@; } </script> "
        Dim sw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(sw)

        Response.Clear()
        Response.Charset = "big5"  ' 在2003EXCEL和2000EXCEL, 中文都不會變亂碼
        Response.ContentType = "Content-Language;content=zh-tw"   ' 新加的
        Response.ContentType = "application/vnd.ms-excel"
        Response.AppendHeader("Content-Disposition", "attachment; filename=Report_14.xls")

        Response.Write("<meta http-equiv=Content-Type content=text/html")
        Response.ContentEncoding = System.Text.UTF8Encoding.UTF8

        GridView1.AllowPaging = False
        GridView1.AllowSorting = False

        'GridView1.Columns.RemoveAt(0)
        'GridView1.Columns.RemoveAt(0)
        'GridView1.Columns.RemoveAt(0)
        'Session("SaveAsExecl_flag") = "1"

        GridView1.DataBind()

        Response.Buffer = True

        title.RenderControl(hw)
        result1.RenderControl(hw)
        GridView1.RenderControl(hw)

        Response.Write(style)
        Response.Write(sw.ToString())
        Response.End()
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        If e.Row.RowIndex <> -1 And Session("SaveAsExecl_flag") = "0" Then

            'CType(e.Row.Cells(0).Controls(0), Button).Attributes.Add("OnClick", "return confirm('Are you sure resend fax?');") '= strJava

            'CType(e.Row.Cells(0).Controls(0), Button).Attributes.Add("OnClick", "if ( !confirm('Are you sure resend fax?')) {return false};")
            'CType(e.Row.Cells(1).Controls(0), Button).Attributes.Add("OnClick", "if ( !confirm('Are you sure resend mail?')) {return false};")
        End If

        If e.Row.RowType = DataControlRowType.DataRow Then
            e.Row.Cells(0).Attributes.Add("class", "text")
            e.Row.Cells(1).Attributes.Add("class", "text")
            e.Row.Cells(2).Attributes.Add("class", "text")
            e.Row.Cells(3).Attributes.Add("class", "text")
           
        End If
        'If e.Row.RowIndex = GridView1.EditIndex And e.Row.RowIndex <> -1 Then
        '    CType(e.Row.Cells(33).Controls(0), TextBox).Width = Unit.Point(100) 'Faxtel
        '    CType(e.Row.Cells(34).Controls(0), TextBox).Width = Unit.Point(20) 'Period
        '    CType(e.Row.Cells(35).Controls(0), TextBox).Width = Unit.Point(20) 'Fax_times
        '    CType(e.Row.Cells(36).Controls(0), TextBox).Width = Unit.Point(20) 'maxtimes
        '    CType(e.Row.Cells(37).Controls(0), TextBox).Width = Unit.Point(50) 'Faxerrcode
        '    CType(e.Row.Cells(38).Controls(0), TextBox).Width = Unit.Point(20) 'EmailFlag
        'End If
    End Sub

    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        '處理'GridView' 的控制項 'GridView' 必須置於有 runat=server 的表單標記之中
        'Dim a = "1"
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""
        '20220217 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc")) '  V6.0   20220215  alex   弱掃 20220215 ; L208 A1 Injection,  Connection String Parameter Pollution
        '20220217 new 
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))) '  V6.0   20220215  alex   弱掃 20220215 ; L208 A1 Injection,  Connection String Parameter Pollution
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
End Class


