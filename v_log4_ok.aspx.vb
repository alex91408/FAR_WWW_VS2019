'
''----------------------------------------------------------------------------------------------
' 每次更改請在此區註解    最新放最上面 alex  p400()  
' ver 要改   #define VERSION XXXX   每次異動版本需要修改
'   ver     date    user      meno
'  ------ -------- --------- -------------------------------------------------------





'  V6.1   20220218  vic    弱掃 20220218

'20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
'Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))






Partial Class v_log4_ok
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Label1.Text = "使用人員：" & Session("acc")

        If Not Me.IsPostBack Then

            '0--------------------------------------------------
            If Session("Login") = False Then
                Response.Redirect("v_login.aspx")
            End If

            lblMDB.Text = Session("v_menu_1")

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

                If set0(iSet) = "v04" Then
                    Exit For
                End If

            Next iSet

            For iSet = 0 To 150
                If set0(iSet) = "v04a" Then
                    SaveAsExecl.Visible = True
                End If

            Next iSet
            '-------------------------------------------------

            Dim sSQL As String
            Dim objReader As Data.OleDb.OleDbDataReader

            '20220223 old Dim mConn As New Data.OleDb.OleDbConnection(Session("ConnectionString"))
            Dim mConn As New Data.OleDb.OleDbConnection(Encoding.UTF8.GetString(Convert.FromBase64String(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString())))

            Dim dv As New Data.DataView
            Dim dt As New Data.DataTable("tabLog")
            Dim row As Data.DataRow
            Dim i, jj As Integer
            Dim sumtrid As Long, sumsucc As Long, sumfail As Long, sumsec As Long

            dv.Table = dt

            dv.Table.Columns.Add("分行")        ' 交易類別")
            dv.Table.Columns.Add("總交易數")        ' 總交易數")
            dv.Table.Columns.Add("成功筆數")        ' 成功筆數")
            dv.Table.Columns.Add("成功筆數(%)")     ' 成功筆數")
            dv.Table.Columns.Add("失敗筆數")        ' 失敗筆數")
            dv.Table.Columns.Add("失敗筆數(%)")     ' 失敗筆數")
            dv.Table.Columns.Add("秒數(秒/通)")  ' 秒數(秒/通)")

            sSQL = " Select BRA01 from [BRA] "

            ' 分行代碼
            '-----------------------------------------------------------
            REM 取db的BRA作表格
            sSQL = ""
            Select Case Session("SAL05")
                Case 1
                    sSQL = " SELECT  BRA01 FROM BRA  where BRA_CHECK_TYPE <> '新增' order by BRA01 "
                Case 2, 3
                    sSQL = " SELECT  BRA01 FROM BRA where BRA01 ='" + Session("SAL03") + "' and  BRA_CHECK_TYPE <> '新增'  order by BRA01 "
                Case 4
                    sSQL = " SELECT  BRA01 FROM BRA where BRA05 ='" + Session("BRA05") + "' and  BRA_CHECK_TYPE <> '新增'  order by BRA01 "
            End Select

            REM 開資料庫
            Dim objcmd0 As New Data.OleDb.OleDbCommand(sSQL, mConn)
            mConn.Open()
            REM 一筆一筆讀出來
            objReader = objcmd0.ExecuteReader
            jj = 0
            Do While objReader.Read
                row = dv.Table.NewRow
                row("分行") = objReader.GetString(0)
                row("總交易數") = "0"
                row("成功筆數") = "0"
                row("失敗筆數") = "0"
                row("秒數(秒/通)") = "0"
                dv.Table.Rows.Add(row)
                jj = jj + 1
            Loop

            '總計row
            row = dv.Table.NewRow
            row("分行") = "total"
            row("總交易數") = "0"
            row("成功筆數") = "0"
            row("失敗筆數") = "0"
            row("秒數(秒/通)") = "0"
            dv.Table.Rows.Add(row)
            mConn.Close()

            '- 呼叫預存程序 
            Dim tSQL As String = ""
            'Using cn As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            Using cn As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))


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

            If Len(tSQL) = "0" Then
                Dim script As String = ""

                script += "<script>"
                script += "alert('查詢日期有誤或查詢不到資料 !!');"
                script += "window.location.href='v_log0_in.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)

            Else

                '-----------------------------------------------------------
                REM 取得交易的秒數

                sSQL = " Select  tt.BRA01,avg(datediff(""s"",t1.stime_in,t1.etime_out)) as avg1 from ( "
                sSQL = sSQL & tSQL
                sSQL = sSQL & " ) as t1,[BRA] tt Where                                                          "
                sSQL = sSQL & " (t1.stime_in >= '" & Session("starttime") & "' and t1.stime_in <= '" & Session("endtime") & "')"
                If Session("SAL05") = 3 Then
                    sSQL = sSQL & " and t1.ao_code ='" & Session("acc") & "'"
                End If
                sSQL = sSQL & "         and ( t1.line_hour is Not null  )      "
                sSQL = sSQL & "         and ( t1.stime_in is Not null  )      "
                sSQL = sSQL & "         and ( t1.etime_out is Not null  )      "
                sSQL = sSQL & "         and ( t1.bra01 = tt.BRA01  )      "
                sSQL = sSQL & " group by tt.BRA01 Order by tt.BRA01"

                REM 開資料庫
                Dim objcmd As New Data.OleDb.OleDbCommand(sSQL, mConn)
                mConn.Open()
                REM 一筆一筆讀出來
                objReader = objcmd.ExecuteReader

                Do While objReader.Read
                    For i = 0 To jj
                        If dv.Table.Rows(i)("分行") = objReader.GetString(0) Then
                            dv.Table.Rows(i)("秒數(秒/通)") = CLng(objReader.GetValue(1))  '& "[" & FormatPercent(objReader.GetInt32(1) / dv.Table.Rows(objReader.GetString(0) - 1)("總交易數"), 2) & "]"
                        End If
                    Next
                Loop

                mConn.Close()

                '-----------------------------------------------------------
                REM 取得成功筆數        
                sSQL = " Select tt.BRA01,count(t1.bra01) as total from (  "
                sSQL = sSQL & tSQL
                sSQL = sSQL & " ) as t1,[BRA] tt Where                                                          "
                sSQL = sSQL & "         (t1.stime_in >= '" & Session("starttime") & "' and t1.stime_in <= '" & Session("endtime") & "')"
                If Session("SAL05") = 3 Then
                    sSQL = sSQL & " and t1.ao_code ='" & Session("acc") & "'"
                End If
                sSQL = sSQL & "         and ( t1.line_hour is Not null  )      "
                sSQL = sSQL & "         and ( t1.stime_in is Not null  )      "
                sSQL = sSQL & "         and ( t1.etime_out is Not null  )      "
                sSQL = sSQL & "         and ( t1.error = '0000' )      "
                sSQL = sSQL & "         and ( t1.bra01 = tt.BRA01  )      "
                sSQL = sSQL & " group by tt.BRA01 Order by tt.BRA01"

                REM 開資料庫
                Dim objcmd2 As New Data.OleDb.OleDbCommand(sSQL, mConn)
                mConn.Open()
                REM 一筆一筆讀出來
                objReader = objcmd2.ExecuteReader

                Do While objReader.Read
                    For i = 0 To jj
                        If dv.Table.Rows(i)("分行") = objReader.GetString(0) Then
                            dv.Table.Rows(i)("成功筆數") = objReader.GetInt32(1)  '& "[" & FormatPercent(objReader.GetInt32(1) / dv.Table.Rows(objReader.GetString(0) - 1)("總交易數"), 2) & "]"
                        End If
                    Next
                Loop

                mConn.Close()

                '-----------------------------------------------------------
                REM 取得失敗筆數          
                sSQL = " Select tt.BRA01,count(t1.bra01) as total from ( "
                sSQL = sSQL & tSQL
                sSQL = sSQL & " ) as t1,[BRA] tt Where "
                sSQL = sSQL & "         (t1.stime_in >= '" & Session("starttime") & "' and t1.stime_in <= '" & Session("endtime") & "')"
                If Session("SAL05") = 3 Then
                    sSQL = sSQL & " and t1.ao_code ='" & Session("acc") & "'"
                End If
                sSQL = sSQL & "         and ( t1.line_hour is Not null  )      "
                sSQL = sSQL & "         and ( t1.stime_in is Not null  )      "
                sSQL = sSQL & "         and ( t1.etime_out is Not null  )      "
                sSQL = sSQL & "         and ( t1.error <> '0000' )      "
                sSQL = sSQL & "         and ( t1.bra01 = tt.BRA01  )      "
                sSQL = sSQL & " group by tt.BRA01 Order by tt.BRA01"

                REM 開資料庫
                Dim objcmd3 As New Data.OleDb.OleDbCommand(sSQL, mConn)
                mConn.Open()
                REM 一筆一筆讀出來
                objReader = objcmd3.ExecuteReader

                Do While objReader.Read
                    For i = 0 To jj
                        If dv.Table.Rows(i)("分行") = objReader.GetString(0) Then
                            dv.Table.Rows(i)("失敗筆數") = objReader.GetInt32(1)  '& "[" & FormatPercent(objReader.GetInt32(1) / dv.Table.Rows(objReader.GetString(0) - 1)("總交易數"), 2) & "]"
                        End If
                    Next
                Loop

                mConn.Close()

                '-----------------------------------------------------------
                '填總計
                For i = 0 To jj - 1
                    dv.Table.Rows(i)("總交易數") = CLng(dv.Table.Rows(i)("成功筆數")) + CLng(dv.Table.Rows(i)("失敗筆數"))
                    sumsucc = sumsucc + CLng(dv.Table.Rows(i)("成功筆數"))
                    sumfail = sumfail + CLng(dv.Table.Rows(i)("失敗筆數"))

                    If CLng(dv.Table.Rows(i)("總交易數")) <> 0 Then
                        dv.Table.Rows(i)("成功筆數") = dv.Table.Rows(i)("成功筆數")
                        dv.Table.Rows(i)("成功筆數(%)") = FormatPercent(dv.Table.Rows(i)("成功筆數") / dv.Table.Rows(i)("總交易數"), 2)
                        dv.Table.Rows(i)("失敗筆數") = dv.Table.Rows(i)("失敗筆數")
                        dv.Table.Rows(i)("失敗筆數(%)") = FormatPercent(dv.Table.Rows(i)("失敗筆數") / dv.Table.Rows(i)("總交易數"), 2)
                    Else
                        dv.Table.Rows(i)("成功筆數") = dv.Table.Rows(i)("成功筆數")
                        dv.Table.Rows(i)("成功筆數(%)") = "0%"
                        dv.Table.Rows(i)("失敗筆數") = dv.Table.Rows(i)("失敗筆數")
                        dv.Table.Rows(i)("失敗筆數(%)") = "0%"
                    End If

                    sumtrid = sumtrid + CLng(dv.Table.Rows(i)("總交易數"))
                    sumsec = sumsec + CLng(dv.Table.Rows(i)("秒數(秒/通)")) * CLng(dv.Table.Rows(i)("總交易數"))
                Next

                dv.Table.Rows(jj)("總交易數") = sumtrid
                If sumtrid <> 0 Then
                    dv.Table.Rows(jj)("成功筆數") = sumsucc
                    dv.Table.Rows(jj)("成功筆數(%)") = FormatPercent(sumsucc / sumtrid, 2)
                    dv.Table.Rows(jj)("失敗筆數") = sumfail
                    dv.Table.Rows(jj)("失敗筆數(%)") = FormatPercent(sumfail / sumtrid, 2)
                    dv.Table.Rows(jj)("秒數(秒/通)") = sumsec \ sumtrid
                Else
                    dv.Table.Rows(jj)("成功筆數") = sumsucc
                    dv.Table.Rows(jj)("成功筆數(%)") = "0%"
                    dv.Table.Rows(jj)("失敗筆數") = sumfail
                    dv.Table.Rows(jj)("失敗筆數(%)") = "0%"
                    dv.Table.Rows(jj)("秒數(秒/通)") = 0
                End If

                '-----------------------------------------------------------
                'gridbind            
                TimeGrid.DataSource = dv.Table
                TimeGrid.DataBind()

                'show result.txt
                result1.Text = "From:" & Session("starttime") & "  -  " & Session("endtime")
                'Response.End()
            End If
        End If
    End Sub

    Private Sub butMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butMenu.Click
        '設定 session 表示使用者已經 login
        Session("Login") = True
        '轉到主選單
        '- 呼叫預存程序 
        Dim tSQL As String = ""
        'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))


            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "4.各分行電話錄音統計表->回主選單"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_log0_in.aspx")
    End Sub

    Private Sub SaveAsExecl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsExecl.Click
        '- 呼叫預存程序 
        Dim tSQL As String = ""



        'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))



            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "4.各分行電話錄音統計表-save execl"

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
        Response.AppendHeader("Content-Disposition", "attachment; filename=Report_04.xls")

        Response.Write("<meta http-equiv=Content-Type content=text/html")
        Response.ContentEncoding = System.Text.UTF8Encoding.UTF8

        Response.Buffer = True

        title.RenderControl(hw)
        result1.RenderControl(hw)
        TimeGrid.RenderControl(hw)

        Response.Write(style)
        Response.Write(sw.ToString())
        Response.End()

    End Sub

    Protected Sub TimeGrid_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TimeGrid.SelectedIndexChanged

    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))

        '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
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
End Class
