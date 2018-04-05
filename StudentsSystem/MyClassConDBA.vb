Imports System.Data
'Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class MyClassConDBA
    'Dim Cn As New SqlClient.SqlConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\Hp\Documents\Visual Studio 2010\Projects\LTCSystem\LTCSystem\bin\Debug\Data\DBTLibya.accdb")
    Dim cmd As New OleDb.OleDbCommand
    Dim tbl As New DataTable
    Dim da As New OleDbDataAdapter
    Dim sql As String
    ' Public conn As New SqlConnection(My.Settings.DBStocksConnectionString)
    'Public TestRunCode As Boolean = True
    ''Insert/Delete/Update

    Private Shared SqlCon As OleDbConnection
    'INSTANT VB WARNING: The following constructor is declared outside of its associated class:
    'ORIGINAL LINE: public DataAccessLayer()

    'Protected Overloads Overrides Function ProcessDialogKey(ByVal keyData As Keys) As Boolean
    '    If keyData = Keys.Enter Then
    '        MyBase.ProcessTabKey(Keys.Tab)
    '        Return True
    '    End If
    '    Return MyBase.ProcessDialogKey(keyData)
    'End Function

    'Protected Overloads Overrides Function ProcessDataGridViewKey(ByVal e As KeyEventArgs) As Boolean
    '    If e.KeyCode = Keys.Enter Then
    '        MyBase.ProcessTabKey(Keys.Tab)
    '        Return True
    '    End If
    '    Return MyBase.ProcessDataGridViewKey(e)
    'End Function


    'دالة تعبئة الداتا قريد 
    Public Function FillDataGrid(ByVal dg As DataGridView, ByVal Sqlstatment As String) As DataSet
        Dim ds As New DataSet
        Dim bs As New BindingSource
        Try
            Dim sda As New OleDbDataAdapter(Sqlstatment, Cn)
            sda.Fill(ds)
            bs.DataSource = ds.Tables(0)
            dg.DataSource = bs
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            Cn.Close()
        End Try
        Return ds
    End Function

    'ربط الا داوت عن طريق DataBind
    Public Sub DataBindObjects(ByVal Perpeties As String, ByVal TbData As DataTable, ByVal FieldName As String, ByVal Obj1 As Object)
        Obj1.DataBindings.clear()
        Obj1.DataBindings.Add(Perpeties, TbData, FieldName)
    End Sub

    Public Sub New()
        SqlCon = New OleDbConnection(My.Settings.DBStudentsConnectionString)
    End Sub

    Public Sub Connect()
        If SqlCon.State <> ConnectionState.Open Then
            SqlCon.Open()
        End If
    End Sub

    Public Sub Disconnect()
        If SqlCon.State <> ConnectionState.Closed Then
            SqlCon.Close()
        End If
    End Sub

    Public Function SelectData(ByVal Stored As String, ByVal Param() As OleDbParameter) As DataTable
        Dim SqlCmd As New OleDbCommand(Stored, SqlCon)
        SqlCmd.CommandType = CommandType.StoredProcedure
        For i As Integer = 0 To Param.Length - 1
            SqlCmd.Parameters.Add(Param(i))
        Next i
        Dim SqlDa As New OleDbDataAdapter(SqlCmd)
        Dim Dt As New DataTable()
        SqlDa.Fill(Dt)
        Return Dt
    End Function

    'Public Sub ExecuteCommand(ByVal Stored As String, ByVal Param() As SqlParameter)
    '    Dim SqlCmd As New SqlCommand(Stored, SqlCon)
    '    SqlCmd.CommandType = CommandType.StoredProcedure
    '    For i As Integer = 0 To Param.Length - 1
    '        SqlCmd.Parameters.Add(Param(i))
    '    Next i
    '    SqlCmd.ExecuteNonQuery()
    'End Sub

    Public Sub AdvanceSearch()
        'Dim sql As String = "SELECT * FROM [table1] "

        'Dim where As String = ""
        'If Me.txtName1.Text.Trim <> "" Then where &= "AND [name1] = @name1 "
        'If Me.txtName2.Text.Trim <> "" Then where &= "AND [name2] = @name2 "
        'If Me.txtName3.Text.Trim <> "" Then where &= "AND [name3] = @name3 "

        'If where <> "" Then sql &= "WHERE " & where.Substring(4)

        'Dim comm As New OleDb.OleDbCommand(sql, conn)
        'If Me.txtName1.Text.Trim <> "" Then comm.Parameters.AddWithValue("@name1", Me.txtName1.Text.Trim)
        'If Me.txtName2.Text.Trim <> "" Then comm.Parameters.AddWithValue("@name2", Me.txtName2.Text.Trim)
        'If Me.txtName3.Text.Trim <> "" Then comm.Parameters.AddWithValue("@name3", Me.txtName3.Text.Trim)

        'MsgBox(sql)
        'Dim dt As New DataTable

        'Dim da As New OleDb.OleDbDataAdapter(comm)

        'If da.Fill(dt) > 0 Then
        '    '
        '    '
        '    '
        'End If
    End Sub
    ''تسجيل عقد جديد للموظف
    'Public Sub NewRegEmp(ByVal TxtNoEmp As Integer, ByVal DateStart As Date, ByVal DateEnd As Date, ByVal IsRunDate As Boolean)
    '    sql = "Select Max(NoStop) from TbEmpStopDesc WHERE (IsDefult) ='True'" ' Order by IsDefult"
    '    Dim NoStop As Int16 = MyClss.ExecuteScalar(sql)
    '    'TbEmploymentMoves
    '    sql = "Select * From TbEmploymentMoves Where Emp_no='" + TxtNoEmp.ToString + "' And IsStop='False'"
    '    Dim Tb As New DataTable
    '    Tb = MyClss.GetRecords(sql)
    '    '  Dim DateStart, DateEnd As DateTime
    '    Dim NoType As Int16 = Tb.Rows(0).Item("NoType")
    '    Dim Tb2 As New DataTable
    '    sql = "Select IsYearOrMonth,Max From TypeEmployment Where NoType='" + NoType.ToString + "'"
    '    Tb2 = MyClss.GetRecords(sql)
    '    Dim IsYearOrMonth As Boolean, Max1 As Int16
    '    If Tb2.Rows.Count > 0 Then
    '        IsYearOrMonth = Tb2.Rows(0).Item("IsYearOrMonth")
    '        Max1 = Tb2.Rows(0).Item("Max")
    '        'Else
    '        '   Exit Sub
    '    End If
    '    Dim Max2 As Int16 = 0
    '    Dim M, D, Y As Int16
    '    Try
    '        If IsRunDate = True Then
    '            DateStart = Tb.Rows(0).Item("DateStart")
    '            DateEnd = Tb.Rows(0).Item("DateEnd")
    '            If IsYearOrMonth = True Then
    '                Max2 = DateDiff(DateInterval.Year, DateStart, DateEnd)
    '                DateStart = DateSerial(Date.Now.Year, DateStart.Month, DateStart.Day)
    '                DateEnd = DateSerial(Date.Now.Year + Max1, DateEnd.Month, DateEnd.Day)
    '            Else
    '                Max2 = DateDiff(DateInterval.Month, DateStart, DateEnd)
    '                If Max2 <> Max1 Then
    '                    M = DateEnd.Month
    '                    Y = DateEnd.Year
    '                    Select Case M
    '                        Case 1, 3, 5, 7, 8, 10, 12
    '                            D = 31
    '                        Case 4, 6, 9, 11
    '                            D = 30
    '                        Case 2
    '                            If Y Mod 4 = 0 Then
    '                                D = 29
    '                            Else
    '                                D = 28
    '                            End If
    '                    End Select

    '                End If
    '            End If
    '        End If
    '        sql = "Update  TbEmploymentMoves set DateStop='" + Format(Date.Now, "yyyy-MM-dd").ToString + "',IsStop='True',NoStop='" + NoStop.ToString + "'"
    '        sql += " WHERE  Emp_no='" + TxtNoEmp.ToString + "' and Id='" + +"'"



    '        MyClss.RunsInsertDeleteUpdateQry(sql)
    '        'MessageBox.Show("تم ابقاف عقد الموظف ", "عملية الحفظ", MessageBoxButtons.OK, MessageBoxIcon.Information)

    '        Dim SQL2 = "INSERT INTO [TbEmploymentMoves] ([Emp_no], [ClassNo], [ActionDate], [SysDate], [Notes], [Dept_No], [Unit_No], "
    '        SQL2 += "[Man_No], [Monaskia_No], [DateStart], [DateEnd], [DateReg], [LevelNo], [NoTypeEmp],[DateFirstWrokFildes],[NoType],[LevelEmpID],[Maktab_No],[IsAutoReg]) VALUES (@Emp_no, "
    '        SQL2 += "@ClassNo, @ActionDate, @SysDate, @Notes, @Dept_No, @Unit_No, @Man_No, @Monaskia_No, @DateStart, "
    '        SQL2 += "@DateEnd, @DateReg, @LevelNo, @NoTypeEmp,@DateFirstWrokFildes,@NoType,@LevelEmpID,@Maktab_No,@IsAutoReg);"


    '        Dim CMD2 As SqlClient.SqlCommand = New SqlClient.SqlCommand
    '        '  If TestAddEmp = True Then
    '        With CMD2
    '            .CommandType = CommandType.Text
    '            .Connection = Cn
    '            .Parameters.Add("@Emp_no", SqlDbType.Int).Value = TxtNoEmp
    '            .Parameters.Add("@ClassNo", SqlDbType.SmallInt).Value = Tb.Rows(0).Item("ClassNo")
    '            .Parameters.Add("@ActionDate", SqlDbType.DateTime).Value = Tb.Rows(0).Item("ActionDate")
    '            .Parameters.Add("@SysDate", SqlDbType.DateTime).Value = DateAndTime.Now.ToString
    '            .Parameters.Add("@Notes", SqlDbType.NVarChar).Value = Tb.Rows(0).Item("Notes")
    '            .Parameters.Add("@Dept_No", SqlDbType.SmallInt).Value = Tb.Rows(0).Item("Dept_No")
    '            .Parameters.Add("@Unit_No", SqlDbType.SmallInt).Value = Tb.Rows(0).Item("Unit_No")
    '            .Parameters.Add("@Man_No", SqlDbType.SmallInt).Value = Tb.Rows(0).Item("Man_No")
    '            .Parameters.Add("@Monaskia_No", SqlDbType.SmallInt).Value = Tb.Rows(0).Item("Monaskia_No")
    '            .Parameters.Add("@Maktab_No", SqlDbType.SmallInt).Value = Tb.Rows(0).Item("Maktab_No")
    '            .Parameters.Add("@DateStart", SqlDbType.DateTime).Value = Tb.Rows(0).Item("DateStart")
    '            .Parameters.Add("@DateEnd", SqlDbType.DateTime).Value = Tb.Rows(0).Item("DateEnd")
    '            .Parameters.Add("@DateReg", SqlDbType.DateTime).Value = Tb.Rows(0).Item("DateReg")
    '            .Parameters.Add("@LevelNo", SqlDbType.Char).Value = Tb.Rows(0).Item("LevelNo")
    '            .Parameters.Add("@NoTypeEmp", SqlDbType.Int).Value = Tb.Rows(0).Item("NoTypeEmp")
    '            .Parameters.Add("@Salary", SqlDbType.Float).Value = 0
    '            .Parameters.Add("@DateFirstWrokFildes", SqlDbType.DateTime).Value = Tb.Rows(0).Item("DateFirstWrokFildes")
    '            .Parameters.Add("@NoType", SqlDbType.SmallInt).Value = Tb.Rows(0).Item("NoType")
    '            .Parameters.Add("@LevelEmpID", SqlDbType.NChar).Value = Tb.Rows(0).Item("LevelEmpID")
    '            .Parameters.Add("@IsAutoReg", SqlDbType.Bit).Value = Tb.Rows(0).Item("IsAutoReg")
    '            .CommandText = SQL2
    '        End With
    '        '  End If
    '        If Cn.State = ConnectionState.Closed Then Cn.Open()

    '        CMD2.ExecuteNonQuery()


    '        sql = "Update  Employment set DateStart='" + Format(DateStart, "yyyy-MM-dd").ToString + "',DateEnd='" + Format(DateEnd, "yyyy-MM-dd").ToString + "'"
    '        sql += " WHERE  Emp_no='" + TxtNoEmp.ToString + "'"



    '        MyClss.RunsInsertDeleteUpdateQry(sql)

    '        '1 New - 2 Insert - 3  Edit  - 4 Del - 5 Print - 6 Find - 7 Login - 8 Logout
    '        MyClss.UserMovement(UserID, "2", TxtNoEmp + "-" + "", "")


    '    Catch ex As SqlException
    '        MsgBox(ex.Message)

    '    End Try

    'End Sub

    'الحصول على بيانات من جدول
    Public Function GetRecords(ByVal Oledb As String) As DataTable
        tbl = New DataTable
        Try
            If Cn.State = ConnectionState.Closed Then Cn.Open()
            da = New OleDbDataAdapter(Oledb, Cn)
            da.Fill(tbl)
        Catch ex As OleDb.OleDbException
            MsgBox(ex.Message)
        Finally
            Cn.Close()
        End Try
        Return tbl
    End Function
    'الحصول عللا البيانات عن طريق الداتاسيت
    Public Function GetRecordsByDataSet(ByVal SqlStr As String, ByVal Oledb As String) As DataSet
        Dim command As New OleDbCommand(SqlStr, Cn)
        Dim adapter As New OleDbDataAdapter(command)
        Dim ds As New DataSet()
        Try
            If Cn.State = ConnectionState.Closed Then Cn.Open()
            adapter.Fill(ds)
            Dim n As Int32 = ds.Tables(0).Rows.Count
        Catch ex As OleDb.OleDbException
            MsgBox(ex.Message)
        Finally
            Cn.Close()
        End Try

        Return ds
    End Function


    Public Function CommandsDisplay(ByVal SqlStr As String, ByVal ColmnName As String) As OleDbDataReader
        Dim command As New OleDbCommand(SqlStr, Cn)
        Dim reader As OleDbDataReader = command.ExecuteReader()
        Try
            If Cn.State = ConnectionState.Open Then Cn.Close()
            Cn.Open()
            While reader.Read()
                Console.WriteLine(reader(ColmnName))
            End While
        Catch ex As OleDb.OleDbException
            MsgBox(ex.Message)
        Finally
            Cn.Close()
        End Try
        Return reader
    End Function

    Public Sub ExecuteCommand(ByVal Sql As String, ByVal Param As OleDbParameter())

        Try
            TestRunCode = True
            Dim SqlCmd As New OleDbCommand(Sql, Cn)
            SqlCmd.CommandType = CommandType.Text
            For i As Int16 = 0 To Param.Length - 1
                SqlCmd.Parameters.Add(Param(i))
            Next
            If Cn.State = ConnectionState.Open Then Cn.Close()
            Cn.Open()
            SqlCmd.ExecuteNonQuery()
        Catch ex As OleDb.OleDbException
            TestRunCode = False
            MsgBox(ex.Message)
        Finally
            Cn.Close()
        End Try
    End Sub
    Public Sub UserMovement(ByVal UserID As Int16, ByVal IDMove As String, ByVal MoveDescrption As String, ByVal MenuName As String)
        Try
            TestRunCode = True

            sql = "INSERT INTO TbUsersMovement( UserID, IDMove, MoveDescrption, MenuName)" 'SysDateStart
            sql += " VALUES     (" + UserID.ToString + "," + IDMove + ",'" + MoveDescrption + "','" + MenuName + "')" '"','" + DateAndTime.Now.ToString +
            MyClss.RunsInsertDeleteUpdateQry(sql)
        Catch ex As OleDb.OleDbException
            TestRunCode = False
            MsgBox(ex.Message)
        Finally

        End Try
    End Sub


    'اجراء الاضافة التعديل والحذف 
    Public Sub RunsInsertDeleteUpdateQry(ByVal Oledb As String)

        Try
            TestRunCode = True
            'conn.ConnectionString = My.Settings.DBHRMConnectionString
            If Cn.State = ConnectionState.Open Then Cn.Close()
            Cn.Open()
            cmd.Connection = Cn
            cmd.CommandText = Oledb
            cmd.ExecuteNonQuery()
        Catch ex As OleDb.OleDbException
            TestRunCode = False
            MsgBox(ex.Message)
        Finally
            Cn.Close()
        End Try

    End Sub

    'اجراء الاضافة التعديل والحذف عن طريق الاجراء المخزن في قاعدة البيانات
    Public Sub RunsInsertUpdateDeleteQry(ByVal StoreProcName As String)

        Try
            TestRunCode = True
            '  conn.ConnectionString = My.Settings.DBHRMConnectionString
            If Cn.State = ConnectionState.Open Then Cn.Close()
            Cn.Open()
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Connection = Cn
            cmd.CommandText = StoreProcName
            cmd.ExecuteNonQuery()
        Catch ex As OleDb.OleDbException
            TestRunCode = False
            MsgBox(ex.Message)
        Finally
            Cn.Close()
        End Try

        'Dim command As New SqlCommand(StoreProcName, conn)
        'command.CommandType = CommandType.StoredProcedure
        'Try
        '    If conn.State = ConnectionState.Open Then conn.Close()
        '    conn.Open()
        '    Dim reader As SqlDataReader = command.ExecuteReader()
        '    While reader.Read()
        '        Console.WriteLine(reader("ColumnName"))
        '    End While
        'Catch ex As SqlException
        '    MsgBox(ex.Message)
        'Finally
        '    conn.Close()
        'End Try
    End Sub

    Public Function ExecuteScalar(ByVal Sql As String) As Object
        Dim command As New OleDbCommand(Sql, Cn)
        Dim number As Integer
        Try
            If Cn.State = ConnectionState.Open Then Cn.Close()
            Cn.Open()
            If Information.IsDBNull(command.ExecuteScalar()) = True Then
                number = 0
            Else
                number = command.ExecuteScalar()
            End If
            ' CInt(command.ExecuteScalar())
        Catch ex As OleDb.OleDbException
            MsgBox(ex.Message)
        Finally
            Cn.Close()
        End Try
        Return number
    End Function

    Public Function ExecuteScalarF(ByVal Sql As String) As Object
        Dim command As New OleDbCommand(Sql, Cn)
        Dim number As Single
        Try
            If Cn.State = ConnectionState.Open Then Cn.Close()
            Cn.Open()
            If Information.IsDBNull(command.ExecuteScalar()) = True Then
                number = 0
            Else
                number = command.ExecuteScalar()
            End If
            ' CInt(command.ExecuteScalar())
        Catch ex As OleDb.OleDbException
            MsgBox(ex.Message)
        Finally
            Cn.Close()
        End Try
        Return number
    End Function
End Class
