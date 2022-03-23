'----------------------------------------------------------------------------------------------
' 每次更改請在此區註解    最新放最上面 CHING2  p400()  
' ver 要改   #define VERSION XXXX   每次異動版本需要修改
'   ver     date    user      meno
'  ------ -------- --------- ------------------------------------------------------------------


'  V6.2   20220223  vic      弱掃 20220223  v_set_d_acc.aspx 內有定義, 要再填新的 解碼過的值 

'Page_Load 加

'20220222 new
'Dim c6 As String
''c6 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))                    'storedproc--> AppSettings
'c6 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.ConnectionStrings("FARConnectionString1").ToString()))     'FARConnectionString1  -> ConnectionStrings
'SqlDataSource6.ConnectionString = c6

'  V6.1   20220216  ching2   弱掃 20220216  
'  V6.0   20220215  ching2   弱掃 20220215


'20220215 old cn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("ConnectionString1").ToString()     '  V6.0   20220215  ching2   弱掃 20220215 ; L338 A1 Injection,  Connection String Parameter Pollution
'         new cn.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("ConnectionString1").ToString()))

'20220215 Old Session("edit_GRP01") = Trim(row.Cells(5).Text)      ' V6.0   20220215  ching2   弱掃 20220215  ; L1109 A1 Injection,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
'20220215 new HttpContext.Current.Session("edit_GRP01") = Trim(row.Cells(5).Text)      ' V6.0   20220215  ching2   弱掃 20220215  ; L1109 A1 Injection,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

'20220216 <Obsolete>
'20220216 old cmdSQL1.CommandText = "SELECT SAL01 FROM SAL WHERE SAL01 = '" & acc1 & "'  and  SAL_CHECK_TYPE <> '新增'  "    '  V6.0   20220215  ching2   弱掃 20220215 ; L55 A1 Injection,  SQL Injection     '  V6.1   20220216  ching2   弱掃 20220216  L72
'20220215 new cmdSQL1.CommandText = "SELECT SAL01 FROM SAL WHERE SAL01 = ? and  SAL_CHECK_TYPE <> '新增'  "    '  V6.0   20220215  ching2   弱掃 20220215 ; L55 A1 Injection,  SQL Injection     '  V6.1   20220216  ching2   弱掃 20220216  L72
'20220215 new cmdSQL1.Parameters.Add("SAL01", acc1)



Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Web
Imports System.Web.SessionState
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports System.Data.OleDb
Imports System.Configuration


Imports System.Security.Cryptography    'MD5


Imports System.Data.SqlClient           ' 970904 ching2 add for  Dim dr As SqlDataReader





Partial Class V_set_d_acc




    Inherits System.Web.UI.Page




    Public edit_idx As Integer

    '  V6.2   20220223  vic      弱掃 20220223
    Public edit_GRP01 As String


    '副程式

    '----------------------------------------------------------------------------------------------------------
    ' funCheck_acc()
    '
    ' 到 SKH.mdf之 acc table 查 acc 是否存在
    '
    '
    '----------------------------------------------------------------------------------------------------------





    <Obsolete>
    Public Function funCheck_GRP(ByVal GRP01 As String) As Boolean



        Dim cnFU1 As System.Data.OleDb.OleDbConnection
        Dim cmdSQL1 As System.Data.OleDb.OleDbCommand

        Dim GRP As String

        cnFU1 = New System.Data.OleDb.OleDbConnection
        cmdSQL1 = New System.Data.OleDb.OleDbCommand

        '20220223 old cnFU1.ConnectionString = Session("ConnectionString")
        cnFU1.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()))

        cmdSQL1.Connection = cnFU1
        cmdSQL1.CommandTimeout = Session("CommandTimeout")

        '20220216 old cmdSQL1.CommandText = "SELECT GRP01 FROM GRP WHERE GRP01 = '" & GRP01 & "'"      ' V6.0   20220215  ching2   弱掃 20220215  ; L70 A1 Injection,  SQL Injection    '  V6.1   20220216  ching2   弱掃 20220216  L86
        cmdSQL1.CommandText = "SELECT GRP01 FROM GRP WHERE GRP01 = ? "      ' V6.0   20220215  ching2   弱掃 20220215  ; L70 A1 Injection,  SQL Injection    '  V6.1   20220216  ching2   弱掃 20220216  L86
        cmdSQL1.Parameters.Add("GRP01", GRP01)

        cnFU1.Open()

        Try
            GRP = CType(cmdSQL1.ExecuteScalar, String)
            If GRP = "" Then
                cnFU1.Close()
                Return False
            Else
                cnFU1.Close()
                Return True
            End If
        Finally

        End Try

        cnFU1.Close()
        Return False              'sql 逾時 

    End Function



    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        'Put user code to initialize the page here

        Label1.Text = "使用人員：" & Session("acc")


        '0--------------------------------------------------
        If Session("Login") = False Then
            Response.Redirect("v_login.aspx")
        End If

        lblMDB.Text = Session("v_menu_1")


        '  V6.2   20220223  vic      弱掃 20220223  v_set_d_acc.aspx 內有定義, 要再填新的 解碼過的值 
        '20220222 new
        Dim c1 As String
        'c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))                    'storedproc--> AppSettings
        c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.ConnectionStrings("FARConnectionString1").ToString()))     'FARConnectionString1  -> ConnectionStrings
        SqlDataSource1.ConnectionString = c1



        addGRP.Enabled = False
        GridView1.Columns(0 + 3).Visible = False
        GridView1.Columns(1 + 3).Visible = False
        GridView1.Columns(0).Visible = False          '覆核
        SaveAsExecl.Visible = False


        'If Not Me.IsPostBack Then


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

            If set0(iSet) = "v13" Then
                Exit For
            End If

        Next iSet
        For iSet = 0 To 150
            If set0(iSet) = "v13a" Then
                SaveAsExecl.Visible = True
            End If
            If set0(iSet) = "v13b" Then
                addGRP.Enabled = True
            End If
            If set0(iSet) = "v13c" Then
                GridView1.Columns(1 + 3).Visible = True
            End If
            If set0(iSet) = "v13d" Then
                GridView1.Columns(0 + 3).Visible = True
            End If

            If set0(iSet) = "v13e" Then
                GridView1.Columns(0).Visible = True
            End If

        Next iSet
        '-------------------------------------------------
        'SqlDataSource1.ConnectionString = Session("ConnectionString")

        lblStat.Text = ""
        lblStat.Visible = False

        Dim sSQL As String = ""

        'ching2 980527 remark sSQL = "select * from GRP where GRP01 <>'" + Session("SAL04") + "' order by GRP01"
        sSQL = "select * from GRP  order by GRP01"
        SqlDataSource1.SelectCommand = sSQL

        'End If

    End Sub

    <Obsolete>
    Private Sub addUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles addGRP.Click

        Dim sSQL As String

        '20220223 old Dim Conn As New OleDb.OleDbConnection(Session("ConnectionString"))
        Dim Conn As New OleDb.OleDbConnection(Encoding.UTF8.GetString(Convert.FromBase64String(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString())))

        Dim objCmd As New OleDbCommand

        '971217 ching2 add for user =""
        '顯示訊息並且轉回首頁

        If Trim(Me.txtGRP01.Text) = "" Or funCheck_GRP(Me.txtGRP01.Text) = True Then

            '- 呼叫預存程序 
            Dim tSQL2 As String = ""

            '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))    ' V6.0   20220215  ching2   弱掃 20220215  ; L187 A1 Injection,  Connection String Parameter Pollution
            '20220215 new
            Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                    Using cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "13.群組管理-群組代碼不可為空白或群組代碼已存在"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using
            End Using

            Dim script1 As String = ""

            script1 += "<script>"
            script1 += "alert('群組代碼不可為空白或群組代碼已存在,請重輸');"
            script1 += "window.location.href='v_set_d_acc.aspx';"
            script1 += "</script>"

            '輸出JavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script1)
            Exit Sub

        End If

        lblStat.Text = ""
        lblStat.Visible = False


        Try
            Dim NDate As String = ""

            NDate = FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 4) & ":" & Second(Now)

            'sSQL = "insert into GRP ( GRP01, GRP02) values  ( '" & txtGRP01.Text & "', '" & txtGRP02.Text & "' )"

            '20220216 old sSQL = "insert into GRP ( GRP_CHECK_TYPE, GRP_CHECK, GRP01, GRP01_, GRP02, GRP02_) values  (  '新增',  'Y', '" & txtGRP01.Text & "', '" & txtGRP01.Text & "', '-',  '" & txtGRP02.Text & "' )"
            sSQL = "insert into GRP ( GRP_CHECK_TYPE, GRP_CHECK, GRP01, GRP01_, GRP02, GRP02_) values  (  '新增',  'Y', ?, ?, '-',  ? )"

            objCmd.CommandText = sSQL      ' V6.0   20220215  ching2   弱掃 20220215  ; L230 A1 Injection,  SQL Injection   '  V6.1   20220216  ching2   弱掃 20220216  L250
            objCmd.Parameters.Add("GRP01", txtGRP01.Text)
            objCmd.Parameters.Add("GRP01_", txtGRP01.Text)
            objCmd.Parameters.Add("GRP02_", txtGRP02.Text)

            objCmd.Connection = Conn
            objCmd.Connection.Open()
            objCmd.ExecuteNonQuery()
            'Response.Write("OK")
        Catch ex As Exception
            'Response.Write(ex.Message)

            '- 呼叫預存程序 

            '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))   ' V6.0   20220215  ching2   弱掃 20220215  ; L240 A1 Injection,  Connection String Parameter Pollution
            '20220215 new
            Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                Using cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "13.群組管理-新增群組代碼失敗(" & txtGRP01.Text & ")"

                    cn1.Open()
                    Dim aaa As Integer = cmd1.ExecuteScalar
                    cmd1.Dispose()
                End Using
            End Using

            lblStat.Text = "新增群組代碼失敗" & ex.Message
            lblStat.Visible = True
            objCmd.Connection.Close()
            Exit Sub

        End Try


        objCmd.Connection.Close()


        '顯示訊息並且轉回首頁
        Dim script As String = ""

        '- 呼叫預存程序 

        '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))     ' V6.0   20220215  ching2   弱掃 20220215  ; L271 A1 Injection,  Connection String Parameter Pollution
        '20220215 new
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))


            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "13.群組管理-新增群組代碼成功(" & txtGRP01.Text & ")"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        script += "<script>"
        script += "alert('新增群組代碼成功" + txtGRP01.Text + " ');"


        script += "window.location.href='v_set_d_acc.aspx';"
        script += "</script>"

        '輸出JavaScript
        ClientScript.RegisterStartupScript(Me.GetType, "", script)


    End Sub


    Protected Sub butMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles butMenu.Click


        '設定 session 表示使用者已經 login
        'Session("Login") = True
        '- 呼叫預存程序 
        Dim tSQL As String = ""

        '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))    ' V6.0   20220215  ching2   弱掃 20220215  ; L309 A1 Injection,  Connection String Parameter Pollution
        '20220215 new
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                Using cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "13.群組管理->回主選單"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using
        End Using

        '轉到主選單
        Response.Redirect("v_log0_in.aspx")

    End Sub

    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        '處理'GridView' 的控制項 'GridView' 必須置於有 runat=server 的表單標記之中
        'Dim a = "1"
    End Sub

    Private Sub SaveAsExecl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsExecl.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))    ' V6.0   20220215  ching2   弱掃 20220215  ; L340 A1 Injection,  Connection String Parameter Pollution
        '20220215 new
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                Using cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "13.群組管理-save excel"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using
        End Using

        GridView1.Columns(0).Visible = False
        GridView1.Columns(3).Visible = False
        GridView1.Columns(4).Visible = False
        Dim ii As Integer

        For ii = 4 To 38
            '980527 ching2 remark 
            'GridView1.Columns(ii).Visible = False
        Next



        Dim style As String = "<style> .text { mso-number-format:\@; } </script> "
        Dim sw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(sw)

        Response.Clear()
        Response.Charset = "big5"  ' 在2003EXCEL和2000EXCEL, 中文都不會變亂碼
        Response.ContentType = "Content-Language;content=zh-tw"   ' 新加的
        Response.ContentType = "application/vnd.ms-excel"
        Response.AppendHeader("Content-Disposition", "attachment; filename=Report_13.xls")

        Response.Write("<meta http-equiv=Content-Type content=text/html")
        Response.ContentEncoding = System.Text.UTF8Encoding.UTF8

        GridView1.AllowPaging = False
        GridView1.AllowSorting = False

        GridView1.DataBind()

        Response.Buffer = True

        'title.RenderControl(hw)

        Dim i As Integer
        Dim j As Integer

        Dim o1 As String


        o1 = ""

        For j = 0 To GridView1.Rows.Count - 1


            txtT.Text = GridView1.Rows(j).Cells(0).Text
            'txtT.RenderControl(hw)

            txtT.Text = GridView1.Rows(j).Cells(1).Text
            'txtT.RenderControl(hw)

            txtT.Text = GridView1.Rows(j).Cells(2).Text
            'txtT.RenderControl(hw)

            txtT.Text = GridView1.Rows(j).Cells(5).Text
            'txtT.RenderControl(hw)

            txtT.Text = GridView1.Rows(j).Cells(6).Text
            'txtT.RenderControl(hw)

            txtT.Text = GridView1.Rows(j).Cells(7).Text
            'txtT.RenderControl(hw)

            txtT.Text = GridView1.Rows(j).Cells(8).Text
            'txtT.RenderControl(hw)



            For i = 9 To GridView1.Columns.Count - 1

                If CType(GridView1.Rows(j).Cells(i).Controls.Item(0), CheckBox).Checked = True Then
                    GridView1.Rows(j).Cells(i).Text = "T"
                    'txtT.Text = "T"
                    'txtT.RenderControl(hw)
                Else
                    GridView1.Rows(j).Cells(i).Text = "F"
                    'txtT.Text = "F"
                    'txtT.RenderControl(hw)
                End If

            Next i

        Next j




        GridView1.RenderControl(hw)

        Response.Write(style)
        Response.Write(sw.ToString())
        Response.End()




        'Dim strIdx As String
        ''顯示訊息並且轉回首頁()
        'Dim script As String = ""

        ''Retrieve the row that contains the button clicked 
        ''by the user from the Rows collection.
        'Dim row As GridViewRow = GridView1.Rows(index)
        'strIdx = row.Cells(7).Text
        'Dim bra_check_type As String








        'Dim ckb As CheckBox

        'ckb = CType(row.Cells(9).Controls.Item(0), CheckBox)          'v01
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(11).Controls.Item(0), CheckBox)          'v01a
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(13).Controls.Item(0), CheckBox)          'v02
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(15).Controls.Item(0), CheckBox)          'v02a
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(17).Controls.Item(0), CheckBox)          'v03
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(19).Controls.Item(0), CheckBox)          'v03a
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(21).Controls.Item(0), CheckBox)          'v04
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(23).Controls.Item(0), CheckBox)          'v04a
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(25).Controls.Item(0), CheckBox)          'v05
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(27).Controls.Item(0), CheckBox)          'v05a
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(29).Controls.Item(0), CheckBox)          'v05b
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If


        'ckb = CType(row.Cells(31).Controls.Item(0), CheckBox)          'v11
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(33).Controls.Item(0), CheckBox)          'v11a
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(35).Controls.Item(0), CheckBox)          'v11b
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(37).Controls.Item(0), CheckBox)          'v11c
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(39).Controls.Item(0), CheckBox)          'v11d
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(41).Controls.Item(0), CheckBox)          'v11e
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If


        'ckb = CType(row.Cells(43).Controls.Item(0), CheckBox)          'v12
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(45).Controls.Item(0), CheckBox)          'v12a
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(47).Controls.Item(0), CheckBox)          'v12b
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(49).Controls.Item(0), CheckBox)          'v12c
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(51).Controls.Item(0), CheckBox)          'v12d
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(53).Controls.Item(0), CheckBox)          'v12e
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(55).Controls.Item(0), CheckBox)          'v12f
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(57).Controls.Item(0), CheckBox)          'v12g
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(59).Controls.Item(0), CheckBox)          'v12h
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(61).Controls.Item(0), CheckBox)          'v12i
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If


        'ckb = CType(row.Cells(63).Controls.Item(0), CheckBox)          'v13
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(65).Controls.Item(0), CheckBox)          'v13a
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(67).Controls.Item(0), CheckBox)          'v13b
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(69).Controls.Item(0), CheckBox)          'v13c
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(71).Controls.Item(0), CheckBox)          'v13d
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(73).Controls.Item(0), CheckBox)          'v13e
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(75).Controls.Item(0), CheckBox)          'v14
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If

        'ckb = CType(row.Cells(77).Controls.Item(0), CheckBox)          'v14a
        'If (ckb.Checked = True) Then
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'Else
        '    txtT.Text = "T"
        '    txtT.RenderControl(hw)
        'End If





    End Sub



    Protected Sub GridView1_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PreRender


        For Each row As GridViewRow In Me.GridView1.Rows
            Session("ac_12") = row.Cells(0 + 5).Text
            CType(row.FindControl("linkbutton3"), LinkButton).Attributes.Add("onclick", "if (confirm('您確定要刪除 " + Session("ac_12") + " 嗎?')==false)  return false;")
        Next


    End Sub


    'Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
    Protected Sub CustomersGridView_RowCommand(ByVal source As Object, ByVal e As GridViewCommandEventArgs)



        If e.CommandName = "Update" Then


            ' 移至   Protected Sub GridView1_RowDeleted( 



            Dim index As Integer    ' = Convert.ToInt32(e.CommandArgument)

            index = Session("edit_idx")    'edit_idx

            Dim strIdx As String
            '顯示訊息並且轉回首頁
            Dim script As String = ""

            ' Retrieve the row that contains the button clicked 
            ' by the user from the Rows collection.
            Dim row As GridViewRow = GridView1.Rows(index)
            strIdx = row.Cells(7).Text
            '20200215 Dim bra_check_type As String
            'bra_check_type = Trim(row.Cells(1).Text)



            'If bra_check_type = "修改" Then

            Try

                Dim SSQL As String = ""

                SSQL = "UPDATE GRP SET GRP_CHECK_TYPE = '修改' , GRP_CHECK = 'Y' "

                SSQL = SSQL & ", GRP01_  ='' , GRP02_  = '" & Trim(row.Cells(7).Text) & "' "



                Dim ckb As CheckBox

                ckb = CType(row.Cells(9).Controls.Item(0), CheckBox)          'v01
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v01_= 'True' "
                Else
                    SSQL = SSQL & ", v01_= 'False'"
                End If

                ckb = CType(row.Cells(11).Controls.Item(0), CheckBox)          'v01a
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v01a_= 'True' "
                Else
                    SSQL = SSQL & ", v01a_= 'False'"
                End If

                ckb = CType(row.Cells(13).Controls.Item(0), CheckBox)          'v02
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v02_= 'True' "
                Else
                    SSQL = SSQL & ", v02_= 'False'"
                End If

                ckb = CType(row.Cells(15).Controls.Item(0), CheckBox)          'v02a
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v02a_= 'True' "
                Else
                    SSQL = SSQL & ", v02a_= 'False' "
                End If

                ckb = CType(row.Cells(17).Controls.Item(0), CheckBox)          'v03
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v03_= 'True' "
                Else
                    SSQL = SSQL & ", v03_= 'False' "
                End If

                ckb = CType(row.Cells(19).Controls.Item(0), CheckBox)          'v03a
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v03a_= 'True' "
                Else
                    SSQL = SSQL & ", v03a_= 'False' "
                End If

                ckb = CType(row.Cells(21).Controls.Item(0), CheckBox)          'v04
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v04_= 'True' "
                Else
                    SSQL = SSQL & ", v04_= 'False' "
                End If

                ckb = CType(row.Cells(23).Controls.Item(0), CheckBox)          'v04a
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v04a_= 'True' "
                Else
                    SSQL = SSQL & ", v04a_= 'False'"
                End If

                ckb = CType(row.Cells(25).Controls.Item(0), CheckBox)          'v05
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v05_= 'True' "
                Else
                    SSQL = SSQL & ", v05_= 'False' "
                End If

                ckb = CType(row.Cells(27).Controls.Item(0), CheckBox)          'v05a
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v05a_= 'True' "
                Else
                    SSQL = SSQL & ", v05a_= 'False'"
                End If

                ckb = CType(row.Cells(29).Controls.Item(0), CheckBox)          'v05b
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v05b_= 'True' "
                Else
                    SSQL = SSQL & ", v05b_= 'False'"
                End If


                ckb = CType(row.Cells(31).Controls.Item(0), CheckBox)          'v11
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v11_= 'True' "
                Else
                    SSQL = SSQL & ", v11_= 'False' "
                End If

                ckb = CType(row.Cells(33).Controls.Item(0), CheckBox)          'v11a
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v11a_= 'True' "
                Else
                    SSQL = SSQL & ", v11a_= 'False'"
                End If

                ckb = CType(row.Cells(35).Controls.Item(0), CheckBox)          'v11b
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v11b_= 'True' "
                Else
                    SSQL = SSQL & ", v11b_= 'False'"
                End If

                ckb = CType(row.Cells(37).Controls.Item(0), CheckBox)          'v11c
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v11c_= 'True' "
                Else
                    SSQL = SSQL & ", v11c_= 'False'"
                End If

                ckb = CType(row.Cells(39).Controls.Item(0), CheckBox)          'v11d
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v11d_= 'True' "
                Else
                    SSQL = SSQL & ", v11d_= 'False'"
                End If

                ckb = CType(row.Cells(41).Controls.Item(0), CheckBox)          'v11e
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v11e_= 'True' "
                Else
                    SSQL = SSQL & ", v11e_= 'False'"
                End If


                ckb = CType(row.Cells(43).Controls.Item(0), CheckBox)          'v12
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v12_= 'True' "
                Else
                    SSQL = SSQL & ", v12_= 'False' "
                End If

                ckb = CType(row.Cells(45).Controls.Item(0), CheckBox)          'v12a
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v12a_= 'True' "
                Else
                    SSQL = SSQL & ", v12a_= 'False'"
                End If

                ckb = CType(row.Cells(47).Controls.Item(0), CheckBox)          'v12b
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v12b_= 'True' "
                Else
                    SSQL = SSQL & ", v12b_= 'False'"
                End If

                ckb = CType(row.Cells(49).Controls.Item(0), CheckBox)          'v12c
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v12c_= 'True' "
                Else
                    SSQL = SSQL & ", v12c_= 'False'"
                End If

                ckb = CType(row.Cells(51).Controls.Item(0), CheckBox)          'v12d
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v12d_= 'True' "
                Else
                    SSQL = SSQL & ", v12d_= 'False'"
                End If

                ckb = CType(row.Cells(53).Controls.Item(0), CheckBox)          'v12e
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v12e_= 'True' "
                Else
                    SSQL = SSQL & ", v12e_= 'False'"
                End If

                ckb = CType(row.Cells(55).Controls.Item(0), CheckBox)          'v12f
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v12f_= 'True' "
                Else
                    SSQL = SSQL & ", v12f_= 'False'"
                End If

                ckb = CType(row.Cells(57).Controls.Item(0), CheckBox)          'v12g
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v12g_= 'True' "
                Else
                    SSQL = SSQL & ", v12g_= 'False'"
                End If

                ckb = CType(row.Cells(59).Controls.Item(0), CheckBox)          'v12h
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v12h_= 'True' "
                Else
                    SSQL = SSQL & ", v12h_= 'False'"
                End If

                ckb = CType(row.Cells(61).Controls.Item(0), CheckBox)          'v12i
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v12i_= 'True' "
                Else
                    SSQL = SSQL & ", v12i_= 'False'"
                End If


                ckb = CType(row.Cells(63).Controls.Item(0), CheckBox)          'v13
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v13_= 'True' "
                Else
                    SSQL = SSQL & ", v13_= 'False' "
                End If

                ckb = CType(row.Cells(65).Controls.Item(0), CheckBox)          'v13a
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v13a_= 'True' "
                Else
                    SSQL = SSQL & ", v13a_= 'False'"
                End If

                ckb = CType(row.Cells(67).Controls.Item(0), CheckBox)          'v13b
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v13b_= 'True' "
                Else
                    SSQL = SSQL & ", v13b_= 'False'"
                End If

                ckb = CType(row.Cells(69).Controls.Item(0), CheckBox)          'v13c
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v13c_= 'True' "
                Else
                    SSQL = SSQL & ", v13c_= 'False'"
                End If

                ckb = CType(row.Cells(71).Controls.Item(0), CheckBox)          'v13d
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v13d_= 'True' "
                Else
                    SSQL = SSQL & ", v13d_= 'False'"
                End If

                ckb = CType(row.Cells(73).Controls.Item(0), CheckBox)          'v13e
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v13e_= 'True' "
                Else
                    SSQL = SSQL & ", v13e_= 'False'"
                End If

                ckb = CType(row.Cells(75).Controls.Item(0), CheckBox)          'v14
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v14_= 'True' "
                Else
                    SSQL = SSQL & ", v14_= 'False' "
                End If

                ckb = CType(row.Cells(77).Controls.Item(0), CheckBox)          'v14a
                If (ckb.Checked = True) Then
                    SSQL = SSQL & ", v14a_= 'True' "
                Else
                    SSQL = SSQL & ", v14a_= 'False'"
                End If







                SSQL = SSQL & " WHERE (GRP01 = '" & Trim(row.Cells(5).Text) & "')"

                '  V6.2   20220223  vic      弱掃 20220223
                '20220215 Old Session("edit_GRP01") = Trim(row.Cells(5).Text)      ' V6.0   20220215  ching2   弱掃 20220215  ; L1109 A1 Injection,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
                '20220215 new 
                '20220216 old HttpContext.Current.Session("edit_GRP01") = Trim(row.Cells(5).Text)      ' V6.0   20220215  ching2   弱掃 20220215  ; L1109 A1 Injection,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session    '  V6.1   20220216  ching2   弱掃 20220216  L1149
                'Dim sTemp5 = Trim(row.Cells(5).Text)
                'HttpContext.Current.Session("edit_GRP01") = sTemp5      ' V6.0   20220215  ching2   弱掃 20220215  ; L1109 A1 Injection,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session    '  V6.1   20220216  ching2   弱掃 20220216  L1149
                '20220223 new
                edit_GRP01 = Trim(row.Cells(5).Text)

                SqlDataSource1.UpdateCommand = SSQL

                SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                SqlDataSource1.Update()


            Catch ex As Exception

                '- 呼叫預存程序 
                Dim tSQL4 As String = ""
                '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))    ' V6.0   20220215  ching2   弱掃 20220215  ; L1121 A1 Injection,  Connection String Parameter Pollution
                '20220215 new
                Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "13.群組管理-修改 失敗(" & Trim(row.Cells(5).Text) & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('修改 失敗,請再試一次" + Trim(row.Cells(5).Text) + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)

            End Try

            'End If



        End If




        If e.CommandName = "Edit" Then

            '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))    ' V6.0   20220215  ching2   弱掃 20220215  ; L1158 A1 Injection,  Connection String Parameter Pollution
            '20220215 new
            Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "13.群組管理-修改"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

        End If



        If e.CommandName = "Cancel" Then
            '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))   ' V6.0   20220215  ching2   弱掃 20220215  ; L1180 A1 Injection,  Connection String Parameter Pollution
            '20220215 new
            Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "13.群組管理-取消修改"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

        End If



        If e.CommandName = "Delete" Then

            '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))    ' V6.0   20220215  ching2   弱掃 20220215  ; L1203 A1 Injection,  Connection String Parameter Pollution
            '20220215 new
            Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "13.群組管理-刪除(" & e.CommandArgument & ")"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using



            Dim index As Integer = Convert.ToInt32(e.CommandArgument)


            Dim SSQL As String = ""

            'SSQL = "delete GRP  WHERE (GRP01 = '" & Trim(row.Cells(5).Text) & "')"
            SSQL = "UPDATE GRP SET GRP_CHECK_TYPE = '刪除' , GRP_CHECK = 'Y' , GRP02_ = '', v01_ = '', v01a_ = '', v02_ = '', v02a_ ='', v03_ ='', v03a_ ='', v04_ ='', v04a_ ='', v05_ ='', v05a_ ='', v05b_ ='', v11_ ='', v11a_ ='', v11b_ ='', v11c_ ='', v11d_ ='', v11e_ ='', v12_ ='', v12a_ ='', v12b_ ='', v12c_ ='', v12d_ ='', v12e_ ='', v12f_ ='', v12g_ ='', v12h_ ='', v12i_ ='', v13_ ='', v13a_ ='', v13b_ ='', v13c_ ='', v13d_ ='', v14_ ='', v14a_ =''  WHERE (GRP01 = '" & Trim(e.CommandArgument) & "')"
            SqlDataSource1.UpdateCommand = SSQL
            SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
            SqlDataSource1.Update()


        End If






        If e.CommandName = "CHECK" Then

            '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))    ' V6.0   20220215  ching2   弱掃 20220215  ; L1243 A1 Injection,  Connection String Parameter Pollution
            '20220215 new
            Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))
                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "13.群組管理-覆核(" & Session("acc") & ")"
                cn1.Open()
                Dim aaa As Integer = cmd1.ExecuteScalar
                cmd1.Dispose()
            End Using




            Dim index As Integer = Convert.ToInt32(e.CommandArgument)
            Dim strIdx As String
            '顯示訊息並且轉回首頁
            Dim script As String = ""

            ' Retrieve the row that contains the button clicked 
            ' by the user from the Rows collection.
            Dim row As GridViewRow = GridView1.Rows(index)
            strIdx = row.Cells(5).Text
            Dim bra_check_type As String
            bra_check_type = Trim(row.Cells(1).Text)



            If bra_check_type = "修改" Then

                Try

                    Dim SSQL As String = ""

                    SSQL = "UPDATE GRP SET GRP_CHECK_TYPE = '-' , GRP_CHECK = 'N' , GRP01_ ='', GRP02_ = '', v01_ = '', v01a_ = '', v02_ = '', v02a_ ='', v03_ ='', v03a_ ='', v04_ ='', v04a_ ='', v05_ ='', v05a_ ='', v05b_ ='', v11_ ='', v11a_ ='', v11b_ ='', v11c_ ='', v11d_ ='', v11e_ ='', v12_ ='', v12a_ ='', v12b_ ='', v12c_ ='', v12d_ ='', v12e_ ='', v12f_ ='', v12g_ ='', v12h_ ='', v12i_ ='', v13_ ='', v13a_ ='', v13b_ ='', v13c_ ='', v13d_ ='', v14_ ='', v14a_ ='' "

                    SSQL = SSQL & ", GRP02  = '" & Trim(row.Cells(8).Text) & "' "


                    '20220215 Dim i As Integer
                    Dim ckb As CheckBox

                    ckb = CType(row.Cells(10).Controls.Item(0), CheckBox)          'v01
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v01= 'True' "
                    Else
                        SSQL = SSQL & ", v01= 'False'"
                    End If

                    ckb = CType(row.Cells(12).Controls.Item(0), CheckBox)          'v01a
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v01a= 'True' "
                    Else
                        SSQL = SSQL & ", v01a= 'False'"
                    End If

                    ckb = CType(row.Cells(14).Controls.Item(0), CheckBox)          'v02
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v02= 'True' "
                    Else
                        SSQL = SSQL & ", v02= 'False'"
                    End If

                    ckb = CType(row.Cells(16).Controls.Item(0), CheckBox)          'v02a
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v02a= 'True' "
                    Else
                        SSQL = SSQL & ", v02a= 'False' "
                    End If

                    ckb = CType(row.Cells(18).Controls.Item(0), CheckBox)          'v03
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v03= 'True' "
                    Else
                        SSQL = SSQL & ", v03= 'False' "
                    End If

                    ckb = CType(row.Cells(20).Controls.Item(0), CheckBox)          'v03a
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v03a= 'True' "
                    Else
                        SSQL = SSQL & ", v03a= 'False' "
                    End If

                    ckb = CType(row.Cells(22).Controls.Item(0), CheckBox)          'v04
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v04= 'True' "
                    Else
                        SSQL = SSQL & ", v04= 'False' "
                    End If

                    ckb = CType(row.Cells(24).Controls.Item(0), CheckBox)          'v04a
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v04a= 'True' "
                    Else
                        SSQL = SSQL & ", v04a= 'False'"
                    End If

                    ckb = CType(row.Cells(26).Controls.Item(0), CheckBox)          'v05
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v05= 'True' "
                    Else
                        SSQL = SSQL & ", v05= 'False' "
                    End If

                    ckb = CType(row.Cells(28).Controls.Item(0), CheckBox)          'v05a
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v05a= 'True' "
                    Else
                        SSQL = SSQL & ", v05a= 'False'"
                    End If

                    ckb = CType(row.Cells(30).Controls.Item(0), CheckBox)          'v05b
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v05b= 'True' "
                    Else
                        SSQL = SSQL & ", v05b= 'False'"
                    End If


                    ckb = CType(row.Cells(32).Controls.Item(0), CheckBox)          'v11
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v11= 'True' "
                    Else
                        SSQL = SSQL & ", v11= 'False' "
                    End If

                    ckb = CType(row.Cells(34).Controls.Item(0), CheckBox)          'v11a
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v11a= 'True' "
                    Else
                        SSQL = SSQL & ", v11a= 'False'"
                    End If

                    ckb = CType(row.Cells(36).Controls.Item(0), CheckBox)          'v11b
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v11b= 'True' "
                    Else
                        SSQL = SSQL & ", v11b= 'False'"
                    End If

                    ckb = CType(row.Cells(38).Controls.Item(0), CheckBox)          'v11c
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v11c= 'True' "
                    Else
                        SSQL = SSQL & ", v11c= 'False'"
                    End If

                    ckb = CType(row.Cells(40).Controls.Item(0), CheckBox)          'v11d
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v11d= 'True' "
                    Else
                        SSQL = SSQL & ", v11d= 'False'"
                    End If

                    ckb = CType(row.Cells(42).Controls.Item(0), CheckBox)          'v11e
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v11e= 'True' "
                    Else
                        SSQL = SSQL & ", v11e= 'False'"
                    End If

                    ckb = CType(row.Cells(44).Controls.Item(0), CheckBox)          'v12
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v12= 'True' "
                    Else
                        SSQL = SSQL & ", v12= 'False' "
                    End If

                    ckb = CType(row.Cells(46).Controls.Item(0), CheckBox)          'v12a
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v12a= 'True' "
                    Else
                        SSQL = SSQL & ", v12a= 'False'"
                    End If

                    ckb = CType(row.Cells(48).Controls.Item(0), CheckBox)          'v12b
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v12b= 'True' "
                    Else
                        SSQL = SSQL & ", v12b= 'False'"
                    End If

                    ckb = CType(row.Cells(50).Controls.Item(0), CheckBox)          'v12c
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v12c= 'True' "
                    Else
                        SSQL = SSQL & ", v12c= 'False'"
                    End If

                    ckb = CType(row.Cells(52).Controls.Item(0), CheckBox)          'v12d
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v12d= 'True' "
                    Else
                        SSQL = SSQL & ", v12d= 'False'"
                    End If

                    ckb = CType(row.Cells(54).Controls.Item(0), CheckBox)          'v12e
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v12e= 'True' "
                    Else
                        SSQL = SSQL & ", v12e= 'False'"
                    End If

                    ckb = CType(row.Cells(56).Controls.Item(0), CheckBox)          'v12f
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v12f= 'True' "
                    Else
                        SSQL = SSQL & ", v12f= 'False'"
                    End If

                    ckb = CType(row.Cells(58).Controls.Item(0), CheckBox)          'v12g
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v12g= 'True' "
                    Else
                        SSQL = SSQL & ", v12g= 'False'"
                    End If

                    ckb = CType(row.Cells(60).Controls.Item(0), CheckBox)          'v12h
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v12h= 'True' "
                    Else
                        SSQL = SSQL & ", v12h= 'False'"
                    End If

                    ckb = CType(row.Cells(62).Controls.Item(0), CheckBox)          'v12i
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v12i= 'True' "
                    Else
                        SSQL = SSQL & ", v12i= 'False'"
                    End If


                    ckb = CType(row.Cells(64).Controls.Item(0), CheckBox)          'v13
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v13= 'True' "
                    Else
                        SSQL = SSQL & ", v13= 'False' "
                    End If

                    ckb = CType(row.Cells(66).Controls.Item(0), CheckBox)          'v13a
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v13a= 'True' "
                    Else
                        SSQL = SSQL & ", v13a= 'False'"
                    End If

                    ckb = CType(row.Cells(68).Controls.Item(0), CheckBox)          'v13b
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v13b= 'True' "
                    Else
                        SSQL = SSQL & ", v13b= 'False'"
                    End If

                    ckb = CType(row.Cells(70).Controls.Item(0), CheckBox)          'v13c
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v13c= 'True' "
                    Else
                        SSQL = SSQL & ", v13c= 'False'"
                    End If

                    ckb = CType(row.Cells(72).Controls.Item(0), CheckBox)          'v13d
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v13d= 'True' "
                    Else
                        SSQL = SSQL & ", v13d= 'False'"
                    End If

                    ckb = CType(row.Cells(74).Controls.Item(0), CheckBox)          'v13e
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v13e= 'True' "
                    Else
                        SSQL = SSQL & ", v13e= 'False'"
                    End If

                    ckb = CType(row.Cells(76).Controls.Item(0), CheckBox)          'v14
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v14= 'True' "
                    Else
                        SSQL = SSQL & ", v14= 'False' "
                    End If

                    ckb = CType(row.Cells(78).Controls.Item(0), CheckBox)          'v14a
                    If (ckb.Checked = True) Then
                        SSQL = SSQL & ", v14a= 'True' "
                    Else
                        SSQL = SSQL & ", v14a= 'False'"
                    End If







                    SSQL = SSQL & " WHERE (GRP01 = '" & Trim(row.Cells(5).Text) & "')"


                    SqlDataSource1.UpdateCommand = SSQL

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Update()

                Catch ex As Exception

                    '- 呼叫預存程序 
                    Dim tSQL4 As String = ""
                    '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))    ' V6.0   20220215  ching2   弱掃 20220215  ; L1551 A1 Injection,  Connection String Parameter Pollution
                    '20220215 new
                    Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                        cmd1.Parameters("@ao_id").Value = Session("acc")
                        cmd1.Parameters("@action").Value = "13.群組管理-覆核-修改 失敗(" & Trim(row.Cells(5).Text) & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('覆核-修改 失敗,請再試一次" + Trim(row.Cells(5).Text) + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '輸出JavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- 呼叫預存程序 
                Dim tSQL3 As String = ""

                '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))    ' V6.0   20220215  ching2   弱掃 20220215  ; L1579 A1 Injection,  Connection String Parameter Pollution
                '20220215 new
                Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "13.群組管理-覆核-修改 成功(" & Trim(row.Cells(5).Text) & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('覆核-修改 成功 " + Trim(row.Cells(5).Text) + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)



            End If


            If bra_check_type = "刪除" Then

                Try

                    Dim SSQL As String = ""

                    SSQL = "delete GRP  WHERE (GRP01 = '" & Trim(row.Cells(5).Text) & "')"


                    Dim a As String = 1
                    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

                    SqlDataSource1.DeleteCommand = SSQL

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Delete()

                Catch ex As Exception

                    '- 呼叫預存程序 
                    Dim tSQL4 As String = ""
                    '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))    ' V6.0   20220215  ching2   弱掃 20220215  ; L1629 A1 Injection,  Connection String Parameter Pollution
                    '20220215 new
                    Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                        cmd1.Parameters("@ao_id").Value = Session("acc")
                        cmd1.Parameters("@action").Value = "13.群組管理-覆核-刪除 失敗(" & Trim(row.Cells(5).Text) & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('覆核-刪除 失敗,請再試一次" + Trim(row.Cells(5).Text) + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '輸出JavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- 呼叫預存程序 
                Dim tSQL3 As String = ""
                '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))    ' V6.0   20220215  ching2   弱掃 20220215  ; L1657 A1 Injection,  Connection String Parameter Pollution
                '20220215 new
                Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "13.群組管理-覆核-刪除 成功(" & Trim(row.Cells(5).Text) & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('覆核-刪除 成功 " + Trim(row.Cells(5).Text) + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)



            End If


            If bra_check_type = "新增" Then

                Try

                    Dim SSQL As String = ""

                    'SSQL = "UPDATE BRA SET BRA01_ = '' , BRA02 = '" & Trim(row.Cells(5).Text) & "' , BRA02_ = '' , BRA03 = '" & Trim(row.Cells(5).Text) & "' , BRA03_ = '' , BRA04 = '" & Trim(row.Cells(7).Text) & "' , BRA04_ = '' , BRA05 = '" & Trim(row.Cells(9).Text) & "', BRA05_ = ''  , BRA_CHECK_TYPE = '-' , BRA_CHECK = 'N' WHERE (BRA01 = '" & Trim(row.Cells(0).Text) & "')"

                    SSQL = "UPDATE GRP SET GRP_CHECK_TYPE = '-' , GRP_CHECK = 'N' , GRP02_ = '', v01_ = '', v01a_ = '', v02_ = '', v02a_ ='', v03_ ='', v03a_ ='', v04_ ='', v04a_ ='', v05_ ='', v05a_ ='', v05b_ ='', v11_ ='', v11a_ ='', v11b_ ='', v11c_ ='', v11d_ ='', v11e_ ='', v12_ ='', v12a_ ='', v12b_ ='', v12c_ ='', v12d_ ='', v12e_ ='', v12f_ ='', v12g_ ='', v12h_ ='', v12i_ ='', v13_ ='', v13a_ ='', v13b_ ='', v13c_ ='', v13d_ ='', v14_ ='', v14a_ ='' "
                    SSQL = SSQL & " WHERE (GRP01 = '" & Trim(row.Cells(5).Text) & "')"


                    Dim a As String = 1
                    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

                    SqlDataSource1.UpdateCommand = SSQL

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Update()

                Catch ex As Exception

                    '- 呼叫預存程序 
                    Dim tSQL4 As String = ""
                    '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))   ' V6.0   20220215  ching2   弱掃 20220215  ; L1710 A1 Injection,  Connection String Parameter Pollution
                    '20220215 new
                    Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                        cmd1.Parameters("@ao_id").Value = Session("acc")
                        cmd1.Parameters("@action").Value = "13.群組管理-覆核-新增 失敗(" & Trim(row.Cells(5).Text) & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('覆核-新增 失敗,請再試一次" + Trim(row.Cells(5).Text) + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '輸出JavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- 呼叫預存程序 
                Dim tSQL3 As String = ""
                '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))    ' V6.0   20220215  ching2   弱掃 20220215  ; L1738 A1 Injection,  Connection String Parameter Pollution
                '20220215 new
                Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "13.群組管理-覆核-新增 成功(" & Trim(row.Cells(5).Text) & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('覆核-新增 成功 " + Trim(row.Cells(5).Text) + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)



            End If


        End If





    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound




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

        End If




    End Sub

    Function FindCommandlink(ByVal Control As Control, ByVal CommandName As String) As LinkButton
        Dim oChildCtrl As Control
        For Each oChildCtrl In Control.Controls
            If TypeOf (oChildCtrl) Is LinkButton Then
                If String.Compare(CType(oChildCtrl, LinkButton).CommandName, CommandName, True) = 0 Then Return oChildCtrl
            End If

        Next
        Return Nothing
    End Function

    Protected Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView1.RowEditing



        edit_idx = e.NewEditIndex

        Session("edit_idx") = edit_idx

    End Sub


    Protected Sub GridView1_RowUpdated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdatedEventArgs) Handles GridView1.RowUpdated

        '20220215 Dim strIdx As String
        '顯示訊息並且轉回首頁

        ' strIdx = Trim(e.Keys.Item(0).ToString)


        Dim SSQL As String = ""

        SSQL = "UPDATE GRP SET GRP_CHECK_TYPE = '修改' , GRP_CHECK = 'Y' "
        SSQL = SSQL & ", GRP02_  = '" & Trim(e.NewValues(0).ToString) & "' "

        '  V6.2   20220223  vic      弱掃 20220223
        '20220223 old SSQL = SSQL & " WHERE (GRP01 = '" & Session("edit_GRP01") & "')"
        SSQL = SSQL & " WHERE (GRP01 = '" & edit_GRP01 & "')"


        SqlDataSource1.UpdateCommand = SSQL
        SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
        SqlDataSource1.Update()


        'Dim strIdx As String
        ''顯示訊息並且轉回首頁

        'strIdx = Trim(e.Keys.Item(0).ToString)


        'Dim SSQL As String = ""

        'SSQL = "UPDATE GRP SET GRP_CHECK_TYPE = '修改' , GRP_CHECK = 'N' "

        'SSQL = SSQL & ", GRP02_  = '" & Trim(e.NewValues(0).ToString) & "' "
        'SSQL = SSQL & ", v01_    = '" & Trim(e.NewValues(1).ToString) & "' "
        'SSQL = SSQL & ", v01a_   = '" & Trim(e.NewValues(2).ToString) & "' "
        'SSQL = SSQL & ", v02_    = '" & Trim(e.NewValues(3).ToString) & "' "
        'SSQL = SSQL & ", v02a_   = '" & Trim(e.NewValues(4).ToString) & "' "
        'SSQL = SSQL & ", v03_    = '" & Trim(e.NewValues(5).ToString) & "' "
        'SSQL = SSQL & ", v03a_   = '" & Trim(e.NewValues(6).ToString) & "' "
        'SSQL = SSQL & ", v04_    = '" & Trim(e.NewValues(7).ToString) & "' "
        'SSQL = SSQL & ", v04a_   = '" & Trim(e.NewValues(8).ToString) & "' "
        'SSQL = SSQL & ", v05_    = '" & Trim(e.NewValues(9).ToString) & "' "
        'SSQL = SSQL & ", v05a_   = '" & Trim(e.NewValues(10).ToString) & "' "
        'SSQL = SSQL & ", v05b_   = '" & Trim(e.NewValues(11).ToString) & "' "
        'SSQL = SSQL & ", v11_    = '" & Trim(e.NewValues(12).ToString) & "' "
        'SSQL = SSQL & ", v11a_   = '" & Trim(e.NewValues(13).ToString) & "' "
        'SSQL = SSQL & ", v11b_   = '" & Trim(e.NewValues(14).ToString) & "' "
        'SSQL = SSQL & ", v11c_   = '" & Trim(e.NewValues(15).ToString) & "' "
        'SSQL = SSQL & ", v11d_   = '" & Trim(e.NewValues(16).ToString) & "' "
        'SSQL = SSQL & ", v11e_   = '" & Trim(e.NewValues(17).ToString) & "' "
        'SSQL = SSQL & ", v12_    = '" & Trim(e.NewValues(18).ToString) & "' "
        'SSQL = SSQL & ", v12a_   = '" & Trim(e.NewValues(19).ToString) & "' "
        'SSQL = SSQL & ", v12b_   = '" & Trim(e.NewValues(20).ToString) & "' "
        'SSQL = SSQL & ", v12c_   = '" & Trim(e.NewValues(21).ToString) & "' "
        'SSQL = SSQL & ", v12d_   = '" & Trim(e.NewValues(22).ToString) & "' "
        'SSQL = SSQL & ", v12e_   = '" & Trim(e.NewValues(23).ToString) & "' "
        'SSQL = SSQL & ", v12f_   = '" & Trim(e.NewValues(24).ToString) & "' "
        'SSQL = SSQL & ", v12g_   = '" & Trim(e.NewValues(25).ToString) & "' "
        'SSQL = SSQL & ", v12h_   = '" & Trim(e.NewValues(26).ToString) & "' "
        'SSQL = SSQL & ", v12i_   = '" & Trim(e.NewValues(27).ToString) & "' "
        'SSQL = SSQL & ", v13_    = '" & Trim(e.NewValues(28).ToString) & "' "
        'SSQL = SSQL & ", v13a_   = '" & Trim(e.NewValues(29).ToString) & "' "
        'SSQL = SSQL & ", v13b_   = '" & Trim(e.NewValues(30).ToString) & "' "
        'SSQL = SSQL & ", v13c_   = '" & Trim(e.NewValues(31).ToString) & "' "
        'SSQL = SSQL & ", v13d_   = '" & Trim(e.NewValues(32).ToString) & "' "
        'SSQL = SSQL & ", v13e_   = '" & Trim(e.NewValues(33).ToString) & "' "
        'SSQL = SSQL & ", v14_    = '" & Trim(e.NewValues(34).ToString) & "' "
        'SSQL = SSQL & ", v14a_   = '" & Trim(e.NewValues(35).ToString) & "' "


        'SSQL = SSQL & " WHERE (GRP01 = '" & Trim(e.Keys.Item(0).ToString) & "')"





        'SqlDataSource1.UpdateCommand = SSQL
        'SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
        'SqlDataSource1.Update()






        '- 呼叫預存程序 
        Dim tSQL As String = ""
        '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))    ' V6.0   20220215  ching2   弱掃 20220215  ; L1908 A1 Injection,  Connection String Parameter Pollution
        '20220215 new
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))


            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
        cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            'cmd1.Parameters("@action").Value = "13.群組管理-編輯群組代碼成功"
            'cmd1.Parameters("@action").Value = "13.群組管理-編輯群組代碼成功(" & Trim(e.OldValues.Item(0).ToString) & "/" & Trim(e.OldValues.Item(1).ToString) & "/" & Trim(e.OldValues.Item(2).ToString) & "/" & Trim(e.OldValues.Item(3).ToString) & "/" & Trim(e.OldValues.Item(4).ToString) & "/" & Trim(e.OldValues.Item(5).ToString) & "->" & Trim(e.NewValues.Item(0).ToString) & "/" & Trim(e.NewValues.Item(1).ToString) & "/" & Trim(e.NewValues.Item(2).ToString) & "/" & Trim(e.NewValues.Item(3).ToString) & "/" & Trim(e.NewValues.Item(4).ToString) & "/" & Trim(e.NewValues.Item(5).ToString) & "/" & ")"

            Dim i As Integer
            Dim o1 As String
            Dim n1 As String


            o1 = ""
            n1 = ""


            o1 = Trim(e.OldValues.Item(4)) & "/"
            n1 = Trim(e.NewValues.Item(0)) & "/"


            For i = 5 To 74 Step 2
                If e.OldValues.Item(i) = True Then
                    o1 = o1 + "T/"
                Else
                    o1 = o1 + "F/"
                End If


            Next i


            For i = 1 To 35

                If e.NewValues.Item(i) = True Then
                    n1 = n1 + "T/"
                Else
                    n1 = n1 + "F/"
                End If

            Next



            cmd1.Parameters("@action").Value = "13.群組管理-編輯群組代碼成功(" & o1 & "->" & n1 & ")"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using




    End Sub




    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click


        '- 呼叫預存程序 
        Dim tSQL As String = ""
        '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))    ' V6.0   20220215  ching2   弱掃 20220215  ; L1977 A1 Injection,  Connection String Parameter Pollution
        '20220215 new
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

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

    Protected Sub GridView1_SelectedIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSelectEventArgs) Handles GridView1.SelectedIndexChanging

        Dim a As Integer
        a = e.NewSelectedIndex

    End Sub
End Class
