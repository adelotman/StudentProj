Imports Microsoft.Win32
Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.IO
Imports DevExpress.XtraBars
Imports DevExpress.XtraBars.Ribbon


Module GenralCodeAccess
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
    Public ConStr As String = "Data Source=" + My.Computer.Name & "\SQLEXPRESS" + ";Initial Catalog=Almohands;Integrated Security=SSPI;persist security info=False;"
    Public SqlConnection1 As New SqlClient.SqlConnection
    Public constring As String
    Dim mcon As New SqlClient.SqlConnection("data source=15.0.0.25\SQLEXPRESS;Initial Catalog=m_admin;integrated security=true;timeout=10")
    Public Cn As New OleDb.OleDbConnection(My.Settings.DBStudentsConnectionString)  'للاتصال والتعرف على قاعدة البيانات
    '  Public Cn As New OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LTCSystems\Data\DBTLibya.accdb")
    Public DAdpt1 As OleDb.OleDbDataAdapter
    Public DAdpt2 As OleDb.OleDbDataAdapter
    Public MyClss As New MyClassConDBA
    Public MyDs As DataSet = New DataSet
    Public MyDsTree As DataSet = New DataSet
    Public GSQLCmd As OleDbCommand = Cn.CreateCommand
    Public GDA As OleDbDataReader = Nothing
    Public GDR As OleDbDataReader = Nothing
    Public DataSource As String
    Public SqlDataAdapter1 As SqlClient.SqlDataAdapter
    Public SqlDataAdapter2 As SqlClient.SqlDataAdapter
    Public SqlDataAdapter3 As SqlClient.SqlDataAdapter
    Public SqlDataAdapter4 As SqlClient.SqlDataAdapter
    Public BUTTONADD As Boolean
    Public BUTTONSAVE As Boolean
    Public BUTTONDELETE As Boolean
    Public BUTTONEDIT As Boolean
    Public BUTTONPRINT As Boolean
    Public SCANFILE As String
    Public ds As DataSet = New DataSet
    Public SqlCMD As SqlClient.SqlCommand = New SqlClient.SqlCommand
    Public USERNAME As String
    Public PASSWORD As String
    Public UserID As Integer
    Public NoEmpLogin As Integer
    Public TxtNoCode As String
    Public SumStops As Int16
    Public TestOpen As Boolean
    Public strReport, SqlRep As String
    Dim cmd As New OleDb.OleDbCommand
    Public TestRunCode As Boolean = True
    Public DataViewName As DataView
    Public beginW As Integer, beginH As Integer
    Public beginRes As String
    Public TimeEndWork(3) As String
    Public DtTakrem As New DataTable
    Public NoEmpl As Int32, NameEmpl As String
    Public myFamilyTable As New DataTable()
    Public NoPassport, TypeStuday, Sql As String



    '-------------------------------------------------------------
    'تحويل البيات الى ميغا
    Public Function FrombytestoMB(ByVal SIZE As Integer) As String
        On Error Resume Next
        Dim result As String = ""
        Select Case SIZE
            Case Is < 1000
                result = Format(Val(SIZE), "#,#0.00") & " Bytes"
            Case Is < 1024000
                result = Format(Val(SIZE / 1024), "#,#0.00") & " KB"
            Case Is < 1024000000
                result = Format(Val(SIZE / 1048576), "#,#0.00") & " MB"
            Case Is >= 1024000000
                result = Format(Val(SIZE / 1073741824), "#,#0.00") & " GB"
        End Select
        FrombytestoMB = result
    End Function
    Sub sendEmail()
        '    Imports System.Net.Mail
        'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '    Dim smtp As New SmtpClient
        '    Dim mail As New MailMessage("myemil@hotmail.com", "emilesend@hotmail.com", TextBox1.Text, TextBox2.Text)
        '    smtp.Credentials = New Net.NetworkCredential("myemil@hotmail.com", "mypassemil")
        '    smtp.EnableSsl = True
        '    smtp.Host = "smtp.live.com"
        'smtp.gmail.com
        'smtp.mail.yahoo.com
        '    smtp.Send(mail)
        'End Sub


    End Sub
    '-------------------------------------------------------------
    Public Sub Empty_Temp()
        On Error Resume Next
        If System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(System.IO.Path.GetTempPath() & "Nasser\")) Then
            System.IO.Directory.Delete(System.IO.Path.GetDirectoryName(System.IO.Path.GetTempPath() & "Nasser\"), True)
        End If
    End Sub

    'دالة تعبئة الداتا قريد 
    Public Function FillDataGrid(ByVal dg As DataGridView, ByVal Sqlstatment As String) As DataSet
        Dim ds As New DataSet
        Dim bs As New BindingSource
        Try
            If Cn.State = ConnectionState.Closed Then Cn.Open()
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

    'ترقيم تسلسل الداتا قريد
    Public Sub SetRowNumber(ByVal DGV As DataGridView)
        For Each row As DataGridViewRow In DGV.Rows
            row.HeaderCell.Value = (row.Index + 1).ToString ' String.Format("{0}", row.Index + 1)
        Next
        DGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
    End Sub

    '------------------ استخدام الاكمال التلقائي في مربع النص ومربع السرد
    Public Sub Autocomplete(ByVal TABLE As String, ByVal FIELD As String, ByVal COMBO As Object)
        Dim Dr As OleDbDataAdapter
        Dim Dt As New DataTable
        Dim str As String = "SELECT DISTINCT " & FIELD & " FROM " & TABLE & " Order By " & FIELD & " Asc"
        Dr = New OleDbDataAdapter(str, Cn)
        Dr.Fill(Dt)
        Dim DataSource As New AutoCompleteStringCollection
        For i As Integer = 0 To Dt.Rows.Count - 1
            DataSource.Add(Dt.Rows(i)(0).ToString)
        Next
        COMBO.AutoCompleteCustomSource = DataSource
        COMBO.AutoCompleteSource = AutoCompleteSource.CustomSource
        COMBO.AutoCompleteMode = AutoCompleteMode.SuggestAppend
    End Sub

    '------------------ استخدام الاكمال التلقائي في مربع النص ومربع السرد
    Public Sub AutocompleteSql(ByVal SQL As String, ByVal COMBO As Object)
        Dim Dr As OleDbDataAdapter
        Dim Dt As New DataTable
        ' Dim str As String = "SELECT DISTINCT " & FIELD & " FROM " & TABLE & " Order By " & FIELD & " Asc"
        Dr = New OleDbDataAdapter(SQL, Cn)
        Dr.Fill(Dt)
        Dim DataSource As New AutoCompleteStringCollection
        For i As Integer = 0 To Dt.Rows.Count - 1
            DataSource.Add(Dt.Rows(i)(0).ToString)
        Next
        COMBO.AutoCompleteCustomSource = DataSource
        COMBO.AutoCompleteSource = AutoCompleteSource.CustomSource
        COMBO.AutoCompleteMode = AutoCompleteMode.SuggestAppend
    End Sub
    Public Sub CLEARDATA1(ByVal frm As Form)
        On Error Resume Next
        Dim txt As Control
        Dim chk As CheckBox
        For Each txt In frm.Controls
            If TypeOf txt Is TextBox Or TypeOf txt Is ComboBox Then
                txt.Text = ""
            End If
        Next
    End Sub

    '--------------   التنقل بالضغط على زر الأدخال   -------------------------
    'لكي يعمل هذا المتغير يجب جعل 
    'خاصية keypreview = True داخل الفورم   
    Public Function ConvertTabKeyToEnterKey(ByVal ev As System.Windows.Forms.KeyEventArgs) As Boolean
        On Error Resume Next
        If ev.KeyCode = Keys.Enter Then
            SendKeys.Send("{tab}")
            ev.SuppressKeyPress = True
        End If
    End Function

    'إدخال التاريخ فقط
    Public Sub PressOnlyDate(ByVal ckeyPressEv As System.Windows.Forms.KeyPressEventArgs) 'As Double
        Dim ckeyPress As Char
        ckeyPress = ckeyPressEv.KeyChar
        If Not (Char.IsDigit(ckeyPress) Or ckeyPress = "-" Or ckeyPress = "/" Or Char.IsControl(ckeyPress)) Then
            ckeyPressEv.Handled = True ' إلغاء الحرف المضغوط
            MsgBox(" لا يسمح بتسجيل الحروف ...  فضلاً إدخل أرقام وإشارات تاريخ فقط", MsgBoxStyle.MsgBoxRight + MsgBoxStyle.Critical, "إدخال خاطئ")
        End If
    End Sub
    '
    ' إدخال أرقام فقط
    Public Sub PressOnlyNumeric(ByVal intKeyValue As System.Windows.Forms.KeyPressEventArgs) ' As Double ' Integer
        Dim chKiy As Char
        chKiy = intKeyValue.KeyChar
        If Not (Char.IsNumber(chKiy) Or chKiy = "." Or Char.IsControl(chKiy)) Then
            MsgBox("غير مسموح بإدخال الحرف .. فضلاً إدخل أرقام فقط", MsgBoxStyle.MsgBoxRight + MsgBoxStyle.Critical, "إدخال خاطئ")
            intKeyValue.Handled = True ' إلغاء الحرف المضغوط
        End If
    End Sub
    'حساب المدة بالايام والاشهر والسنة
    Public Function TakremCalc(ByVal dteStartDate As DateTime, ByVal dteEndDate As DateTime) As String
        'Dim dteStartDate As Date = #12/9/1972#
        'Dim dteEndDate As Date = #7/30/2013#

        Dim dteResultYears As Long = DateDiff(DateInterval.Year, dteStartDate, dteEndDate)
        Dim helpDate As Date = dteStartDate.AddYears(Convert.ToInt32(dteResultYears))

        If helpDate > dteEndDate Then
            dteResultYears -= 1
            dteStartDate = dteStartDate.AddYears(Convert.ToInt32(dteResultYears))
        Else
            dteStartDate = helpDate
        End If

        Dim dteResultMonths As Long = DateDiff(DateInterval.Month, dteStartDate, dteEndDate)
        helpDate = dteStartDate.AddMonths(Convert.ToInt32(dteResultMonths))

        If helpDate > dteEndDate Then
            dteResultMonths -= 1
            dteStartDate = dteStartDate.AddMonths(Convert.ToInt32(dteResultMonths))
        Else
            dteStartDate = helpDate
        End If

        Dim dteResultDays As Long = DateDiff(DateInterval.Day, dteStartDate, dteEndDate)

        helpDate = dteStartDate.AddDays(Convert.ToInt32(dteResultDays))

        If helpDate > dteEndDate Then
            dteResultDays -= 1
            dteStartDate = dteStartDate.AddDays(Convert.ToInt32(dteResultDays))
        Else
            dteStartDate = helpDate
        End If
        Dim sb As New System.Text.StringBuilder
        Dim ageSTR(3) As String

        TimeEndWork(0) = dteResultYears.ToString
        TimeEndWork(1) = dteResultMonths.ToString
        TimeEndWork(2) = dteResultDays.ToString()
        'sb.AppendLine("Difference:")
        'sb.AppendLine("Years: " & dteResultYears.ToString)
        'sb.AppendLine("Months: " & dteResultMonths.ToString)
        'sb.AppendLine("Days: " & dteResultDays.ToString)
        '  MsgBox(sb.ToString)
        Return TimeEndWork(3)
    End Function




    'حساب المدة بالايام والاشهر والسنة
    Public Function AgeCalc(ByVal DatePast As DateTime, _
                              Optional ByVal DateFuture As DateTime = Nothing, _
                              Optional ByVal returnMonthsAndDays As Boolean = False) As String

        'calculate age using two dates. 
        'the age returned will be years and days or years months days
        'Usage:
        'Dim ageSTR() As String = AgeCalc(#2/29/1008#, #2/28/2009#).Split("|"c)

        'Debug.WriteLine(ageSTR(0)) 'years
        'Debug.WriteLine(ageSTR(1)) 'days
        'Debug.WriteLine(ageSTR(2)) 'how long calculation took (in ticks)
        'or
        'Debug.WriteLine(ageSTR(0)) 'years
        'Debug.WriteLine(ageSTR(1)) 'months
        'Debug.WriteLine(ageSTR(2)) 'days
        'Debug.WriteLine(ageSTR(3)) 'how long calculation took (in ticks)
        '

        'note - the wonderful thing about DateSerial is that it takes care of 
        'Leap Birthdays AKA leaplings
        'a leaplings b-day is 3/1 for non-leap years
        Dim retval As String
        Dim ts As New TimeSpan
        Dim dp, df, nxtAnniv As DateTime
        Dim years, days As Integer, stpw As New Stopwatch
        stpw.Start() 'provide metric
        If DateFuture = Nothing Then DateFuture = DateTime.Now 'use now for future date
        If DatePast > DateFuture Then 'make sure the past is the past
            dp = DateFuture
            df = DatePast
        Else
            dp = DatePast
            df = DateFuture
        End If
        years = df.Year - dp.Year 'calculate years
        'create next "birthday"
        'nxtAnniv = New System.DateTime(df.Year, dp.Month, dp.Day) 'does not take care of leap b-day
        nxtAnniv = DateSerial(df.Year, dp.Month, dp.Day) 'takes care of leap b-day
        If nxtAnniv > df Then 'is the next "birthday" > future date
            years -= 1 'yes, adjust year
            nxtAnniv = DateSerial(df.Year - 1, dp.Month, dp.Day) 're-create so it is before future
        End If
        ts = df - nxtAnniv 'will give days
        days = ts.Days
        Dim mos As Integer = 0
        If returnMonthsAndDays Then 'return years mos days
            Dim dfmonths As Long
            dfmonths = DateDiff(DateInterval.Month, nxtAnniv, df)
            df = df.Date
            Dim na As DateTime = nxtAnniv.Date 'get anniversary date
            Dim la As DateTime = nxtAnniv.Date 'set trailing anniversary
            Do While na < df
                na = na.AddMonths(1) 'add one month
                If na < df Then 'less than future date
                    Dim td As Integer = (na - la).Days
                    If td <= days Then 'enough days?
                        days -= td 'subtract days
                        mos += 1 'increment months
                        la = na 'set trailing anniversary
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop
            retval = years.ToString & "|" & mos.ToString & "|" & days.ToString   'create return value
        Else 'return years and days
            retval = years.ToString & "|" & days.ToString   'create return value
            stpw.Stop() 'how long did it take?
        End If
        retval &= "|" & stpw.ElapsedTicks.ToString
        ' MsgBox(retval.ToString)
        Return retval
    End Function




    Public Sub FormsDefultSetting(ByVal frm As Form)
        With frm
            .StartPosition = FormStartPosition.CenterScreen
            .RightToLeft = RightToLeft.Yes
            .RightToLeftLayout = True
            .FormBorderStyle = FormBorderStyle.Fixed3D
            .Font = New Font("Tahoma", 9, FontStyle.Bold)

        End With
    End Sub


    'تخصيي الاختيار للاداة تشك ليست بوكس
    Public Sub SetChecked(ByVal clb As CheckedListBox, ByVal _
    check_items As Boolean)
        For i As Integer = 0 To clb.Items.Count - 1
            clb.SetItemChecked(i, check_items)
        Next i
    End Sub

    Public Sub TreeviewDisplay(ByVal Dataset1 As DataSet, ByVal Table1 As String, ByVal FiledDisplay As String, ByVal FiledId As String, ByVal TreeViewName As TreeView)
        Try
            'حذف كل العقد الموجودة سابقاً
            TreeViewName.Nodes.Clear()
            'التصريح عن داتا فيو لعرض البيانات من خلالها ,, ويمكن اسختدام ال
            'BindingSource

            DataViewName = Dataset1.Tables(0).DefaultView
            'تعين فلتر للداتا فيو للحصول فقط على العقد الأباء
            DataViewName.RowFilter = "ParentID ='0'"
            'تعليمة تكرارية لإضافة العقد الأباء إلى التري فيو
            For Each a As DataRowView In DataViewName
                'التصريح عن عقدة جديدة
                Dim trn As New TreeNode(a(FiledDisplay))
                trn.Tag = a(FiledId).ToString
                Dim co As New ColorConverter
                Dim color1 As New Color
                '  color1 = co.ConvertFromString(a("Color").ToString)
                ' trn.ForeColor = color1
                'إضافة العقدة إلى التري فيو
                TreeViewName.Nodes.Add(trn)
            Next
            'تحميل العقد الأبناء
            For Each node As TreeNode In TreeViewName.Nodes
                AddSubNodeSub(node, Dataset1, Table1, FiledDisplay, FiledId)
            Next

            TreeViewName.ExpandAll()
            TreeViewName.Select()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Sub AddSubNodeSub(ByVal Node As TreeNode, ByRef Data1 As DataSet, ByVal tab As String, ByVal FiledDisplay As String, ByVal FiledId As String)
        Dim Datav1 As DataView = Data1.Tables(0).DefaultView
        Datav1.RowFilter = "ParentID =" & Val(Node.Tag)
        For Each a As DataRowView In Datav1
            Dim SubNode As New TreeNode(a(FiledDisplay).ToString())
            SubNode.Tag = a(FiledId).ToString
            Dim co As New ColorConverter
            Dim color1 As New Color
            ' color1 = co.ConvertFromString(a("Color").ToString)
            '  SubNode.ForeColor = color1
            Node.Nodes.Add(SubNode)
            If Not SubNode.Tag Is String.Empty Then
                AddSubNodeSub(SubNode, MyDsTree, tab, FiledDisplay, FiledId)
            End If
        Next

    End Sub

    Public Sub ConData()
        'الاتصال بقاعدة البيانات
        Dim Str As String = "Data Source=.\SQLEXPRESS;AttachDbFilename=C:\Almohands\Database\DbAlmohands.mdf;" & _
           "Integrated Security=True;Connect Timeout=30;User Instance=True"

        'If GETATTACHDATABASENAME() = False Then
        'ATTACHDATABASENAME("Almohands", Application.StartupPath & "\DATABASE\Almohands.MDF", Application.StartupPath & "\DATABASE\Almohands_log.LDF")
        'End If
        Str = "Data Source=.\SQLEXPRESS;AttachDbFilename=" & Application.StartupPath & "\Database\DbAlmohands.mdf;" & _
           "Integrated Security=True;Connect Timeout=30;User Instance=True"
        ' MsgBox(Str)
        Cn.Close()
        Try
            Cn.ConnectionString = My.Settings.DBStudentsConnectionString  ' Str
            Cn.Open()
            If Cn.State = ConnectionState.Open Then
                '  MsgBox("Opened")
            Else
                ' MsgBox("Closed")
            End If
        Catch ex As SqlClient.SqlException ' Exception
            MsgBox(ex.ToString)
        End Try


    End Sub
    'بحث عن بيانات بدلالة حقل واحد
    Public Function FindNoCode(ByVal TABLE As String, ByVal FIELD1 As String, ByVal StrFind As String) As Boolean
        On Error Resume Next
        Dim DS As New DataSet
        ' Dim str1 As String = ""
        Dim str As String = "SELECT Count(*) From " & TABLE & " Where " & FIELD1 & "='" + StrFind.ToString + "'"
        Dim Sql2 As String
        If Cn.State = ConnectionState.Closed Then Cn.Open()
        Dim ADP As New OleDbDataAdapter(str, Cn)

        DS.Clear()
        ADP.Fill(DS, "Tb1")
        '  Dim Test = False
        FindNoCode = False
        If DS.Tables(0).Rows(0).Item(0) > 0 Then FindNoCode = True
        '  MessageBox.Show("اسف .. لم يتم العثور على اي نتيجة بحث .. تاكد من ادخال نص البحث", "فشل البحث ", MessageBoxButtons.OK, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
        '  Exit Function

        '  End If
    End Function

    Public Function GETATTACHDATABASENAME() As Boolean
        On Error Resume Next
        Dim DS As New DataSet
        Dim SqlConnection1 As SqlClient.SqlConnection = New SqlClient.SqlConnection("Data Source=" + My.Computer.Name & "\SQLEXPRESS" + ";Initial Catalog=tempdb;Integrated Security=SSPI;")
        Dim str As String = "Select DISTINCT name from master.dbo.sysdatabases where name Like 'Almohands' and has_dbaccess(Name) = 1 "
        Dim ADP As SqlClient.SqlDataAdapter
        ADP = New SqlClient.SqlDataAdapter(str, SqlConnection1)
        DS.Clear()
        ADP.Fill(DS)
        Dim i As Integer
        If DS.Tables(0).Rows.Count = 0 Then
            GETATTACHDATABASENAME = False
            MessageBox.Show(" قاعدة البيانات  'Almohands'" & "غير متصلة بالسرفر جارى عمل الاتصال" & vbCrLf & vbCrLf & " من فضلك انتظر قليلاَ                             ", My.Computer.Name, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading)
        Else
            GETATTACHDATABASENAME = True
        End If
        ADP.Dispose()
        SqlConnection1.Dispose()
    End Function
    Public Sub ATTACHDATABASENAME(ByVal MYDBNAME As String, ByVal f1lepathprimary As String, ByVal f1lepathlog As String)
        Try
            Dim SqlConnection1 As SqlClient.SqlConnection = New SqlClient.SqlConnection("Data Source=" + My.Computer.Name & "\SQLEXPRESS" + ";Initial Catalog=tempdb;Integrated Security=SSPI;")
            Dim CMD As SqlClient.SqlCommand = New SqlClient.SqlCommand
            CMD.CommandType = CommandType.Text
            CMD.Connection = SqlConnection1
            If SqlConnection1.State = ConnectionState.Open Then SqlConnection1.Close()
            SqlConnection1.Open()
            CMD.CommandText = "sp_attach_db " & MYDBNAME & ",'" & f1lepathprimary & "'" & ",'" & f1lepathlog & "'"
            ' CMD.CommandText = "CREATE DATABASE " & MYDBNAME & " ON (FILENAME = '" & f1lepathprimary & "')FOR ATTACH"
            CMD.ExecuteNonQuery()
            SqlConnection1.Dispose()
            MessageBox.Show("تم انشاء اتصال  قاعدة البيانات 'Almohands' بالسرفر ", "ATTCH DATABASE", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
        Catch ex As Exception
            Dim result As Integer
            result = MessageBox.Show("فشل البرنامج فى انشاء اتصال  قاعدة البيانات 'Almohands' بالسرفر", "ATTCH DATABASE", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
        End Try
    End Sub

    Public Sub FILLCOMBOBOXITEMS1(ByVal TABLE As String, ByVal FIELD1 As String, ByVal FIELD2 As String, ByVal COMBO As ComboBox)
        'On Error Resume Next

        Try
            'Dim I As Integer
            Dim DS As New DataSet
            Dim str As String = "SELECT  " & FIELD1 & "," & FIELD2 & " FROM " & TABLE & " Order By " & FIELD1 & " Asc"
            If Cn.State = ConnectionState.Closed Then Cn.Open()
            Dim ADP As New OleDb.OleDbDataAdapter(str, Cn)
            DS.Clear()
            ADP.Fill(DS, "TBL")
            '  COMBO.Items.Clear()
            COMBO.DataSource = DS.Tables(0)
            COMBO.DisplayMember = DS.Tables(0).Columns(1).ColumnName
            COMBO.ValueMember = DS.Tables(0).Columns(0).ColumnName
            ADP.Dispose()
            Cn.Close()
        Catch ex As SqlClient.SqlException
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Sub FILLCOMBOBOXITEMS2(ByVal Sql As String, ByVal COMBO As Object)
        'On Error Resume Next

        Try
            'Dim I As Integer
            Dim DS As New DataSet
            Dim str As String = Sql
            If Cn.State = ConnectionState.Closed Then Cn.Open()
            Dim ADP As New OleDb.OleDbDataAdapter(str, Cn)
            DS.Clear()
            ADP.Fill(DS, "TBL")
            '  COMBO.Items.Clear()
            COMBO.DataSource = DS.Tables(0)
            COMBO.ValueMember = DS.Tables(0).Columns(0).ColumnName
            COMBO.DisplayMember = DS.Tables(0).Columns(1).ColumnName

            ADP.Dispose()
            Cn.Close()
        Catch ex As SqlClient.SqlException
            MsgBox(ex.Message)
        End Try

    End Sub
    Public Sub FILLCOMBOBOXITEMS(ByVal TABLE As String, ByVal FIELD As String, ByVal COMBO As Object)
        'On Error Resume Next
        Dim I As Integer
        Dim DS As New DataSet
        Dim str As String = "SELECT DISTINCT " & FIELD & " FROM " & TABLE & " Order By " & FIELD & " Asc"
        If Cn.State = ConnectionState.Closed Then Cn.Open()
        Dim ADP As New OleDb.OleDbDataAdapter(str, Cn)
        DS.Clear()
        ADP.Fill(DS, "TBL")
        COMBO.Items.Clear()
        For I = 0 To DS.Tables("TBL").Rows.Count - 1
            COMBO.Items.Add(DS.Tables("TBL").Rows(I).Item(0))
        Next I
        ADP.Dispose()
    End Sub


    Public Sub MYSHUTDOWN(ByVal OPERATION As Byte)
        On Error Resume Next
        Dim resault As Integer
        Select Case OPERATION
            Case 0
                resault = MessageBox.Show("هل تريد غلق الجهاز", "اغلاق الجهاز ", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
                If resault = vbYes Then
                    Shell("shutdown -s -f -t 0")
                Else
                    Exit Sub
                End If
            Case 1
                resault = MessageBox.Show("هل تريد اعادة تشغيل الويندز ", "اعادة تشغيل الويندز ", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
                If resault = vbYes Then
                    Shell("shutdown -r -f -t 0")
                Else
                    Exit Sub
                End If
            Case 2
                resault = MessageBox.Show("هل تريد عمل Log Off ", "Log Off ", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
                If resault = vbYes Then
                    Shell("shutdown -l -f -t 0")
                Else
                    Exit Sub
                End If
        End Select
    End Sub
    'حذف سجل بدلالة حقل واحد
    Public Sub MYDELETERECORD(ByVal TABLE As String, ByVal FIELD As String, ByVal TXT As Object, ByVal BS As BindingSource, Optional ByVal FIELDtextornumer As Boolean = True)
        On Error Resume Next
        Dim resault As Integer
        Dim SQL As String = ""
        If FIELDtextornumer = True Then
            SQL = "DELETE FROM " & TABLE & " WHERE " & FIELD & "=" & TXT.Text.Trim
        Else
            SQL = "DELETE FROM " & TABLE & " WHERE " & FIELD & "='" & TXT.Text.Trim & "'"
        End If

        If (BS.Count > 0) Then
            resault = MessageBox.Show(" سيتم حذف الصف الحالى", "حذف سجل " & TABLE, MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
            If resault = vbYes Then
                Dim CMD As SqlClient.SqlCommand = New SqlClient.SqlCommand
                CMD.CommandType = CommandType.Text
                CMD.Connection = SqlConnection1
                If SqlConnection1.State = ConnectionState.Open Then SqlConnection1.Close()
                SqlConnection1.Open()
                CMD.CommandText = SQL
                CMD.ExecuteNonQuery()
                SqlConnection1.Close()
            Else
                MessageBox.Show("تم ايقاف عملية الحذف", "حذف سجل", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
                Exit Sub
            End If
        Else
            MessageBox.Show(" لايوجد سجل حالى لاتمام عملية الحذف", "حذف سجل", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
            Exit Sub
        End If
    End Sub
    'حذف سجل بدلالة حقلين
    Public Sub DeleteRec(ByVal TABLE As String, ByVal FIELD1 As String, ByVal FIELD2 As String, ByVal TXT1 As Object, ByVal TXT2 As Object)
        '  On Error Resume Next
        Dim resault As Integer
        Dim SQL As String

        SQL = "DELETE FROM " & TABLE.Trim & " WHERE " & FIELD1.Trim & "='" & TXT1.ToString & "' And " & FIELD2.Trim & "='" & TXT2.ToString & "'"
        'Sql="Delete From " & Table & " Where " & field & "='" & txt.
        resault = MessageBox.Show("هل انت متأكد من عملية حذف بيانات الصف الحالي  ؟", "رسالة حذف " & TABLE, MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
        If resault = vbYes Then
            Dim CMD As OleDbCommand = New OleDbCommand
            CMD.CommandType = CommandType.Text
            CMD.Connection = Cn
            If Cn.State = ConnectionState.Open Then Cn.Close()
            Cn.Open()
            CMD.CommandText = SQL
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            Cn.Close()
            MessageBox.Show(" لقد تم حذف بيانات  بنجاح", "نجاح عملية الحذف", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
        Else
            MessageBox.Show("تم ايقاف عملية الحذف", "حذف سجل", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
            Exit Sub
        End If

    End Sub

    'حذف جميع السجلات
    Public Sub MYDELETEALL2(ByVal TABLE As String, ByVal BS As BindingSource)
        On Error Resume Next
        Dim resault As Integer
        Dim SQL As String = ""
        SQL = "DELETE FROM " & TABLE
        If (BS.Count > 0) Then
            resault = MessageBox.Show("سيتم حذف جميع البيانات من الجدول الحالي .. هل انت متأكد ؟", "حذف الكل " & TABLE, MessageBoxButtons.YesNo, MessageBoxIcon.Stop, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
            If resault = vbYes Then
                Dim CMD As SqlClient.SqlCommand = New SqlClient.SqlCommand
                CMD.CommandType = CommandType.Text
                CMD.Connection = SqlConnection1
                If SqlConnection1.State = ConnectionState.Open Then SqlConnection1.Close()
                SqlConnection1.Open()
                CMD.CommandText = SQL
                CMD.ExecuteNonQuery()
                SqlConnection1.Close()
            Else
                MessageBox.Show("تم ايقاف عملية الحذف", "حذف الكل", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
                Exit Sub
            End If
        Else
            MessageBox.Show(" لايوجد سجلات  لاتمام عملية الحذف", "حذف الكل", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading)
            Exit Sub
        End If
    End Sub
    'غلق وفتح القوائم الرئيسية 
    Public Sub CloseOpenBtnMenue(ByVal TypeOpen As Int16)
        Dim mCurrentItem As BarItem
        Dim mBarSubItem As BarSubItem
        '  Dim mSubLink As BarItemLink
        Dim RecId As Int16 = 0
        Dim Sql1, Sql2 As String
        '= "Select Count(*) From TbMenues where ButtonName='"
        Dim SqlStr As String = ""
        Dim DsTest As New DataSet
        '  Dim TestBtn As Boolean
        Sql1 = "Select IDMenue from TbMenues where ButtonName='"
        Sql2 = "Select IsLogin from TbViewUsersAllowMenus where UserID =" + UserID.ToString + " and ButtonName='"
        For Each currentPage As RibbonPage In FrmMainMenue.RibbonControl.Pages

            'DsTest.Clear()
            'SqlStr = Sql + currentPage.Name + "'"
            'DsTest = MyClss.GetRecordsByDataSet(SqlStr, "")
            'If DsTest.Tables(0).Rows.Count = 0 Then
            '    IdRef1 = DsTest.Tables(0).Rows(0).Item(0)
            'End If

            For Each currentGroup As RibbonPageGroup In currentPage.Groups
                'DsTest.Clear()
                'SqlStr = Sql + currentGroup.Name + "'"
                'DsTest = MyClss.GetRecordsByDataSet(SqlStr, "")
                'If DsTest.Tables(0).Rows.Count = 0 Then
                '    IdRef2 = DsTest.Tables(0).Rows(0).Item(0)
                'End If
                For Each currenLink As BarItemLink In currentGroup.ItemLinks
                    mCurrentItem = currenLink.Item
                    DsTest.Clear()
                    If TypeOpen = 1 Then
                        SqlStr = Sql1 + mCurrentItem.Name + "'"
                    ElseIf TypeOpen = 2 Then
                        SqlStr = Sql2 + mCurrentItem.Name + "'"
                    ElseIf TypeOpen = 3 Then
                        mCurrentItem.Enabled = True
                    End If
                    DsTest = MyClss.GetRecordsByDataSet(SqlStr, "")
                    If DsTest.Tables(0).Rows.Count > 0 Then
                        If TypeOpen = 1 Then
                            mCurrentItem.Enabled = False
                        ElseIf TypeOpen = 2 Then
                            mCurrentItem.Enabled = DsTest.Tables(0).Rows(0).Item(0)
                        End If
                    End If
                    If mCurrentItem.GetType.FullName = "DevExpress.XtraBars.BarSubItem" Then
                        mBarSubItem = mCurrentItem
                        'For Each mSubLink In mBarSubItem.ItemLinks
                        '    'DsTest.Clear()
                        '    'SqlStr = Sql + mSubLink.Item.Name + "'"
                        '    'DsTest = MyClss.GetRecordsByDataSet(SqlStr, "")
                        '    'If DsTest.Tables(0).Rows.Count = 0 Then
                        '    '    RecId = MyClss.ExecuteScalar("Select Max(IDMenue) From TbMenues ") + 1
                        '    'End If
                        '    'Debug.Print(currentPage.Text & " - " & currentGroup.Text & " - " & mSubLink.Item.Caption & " - " & mSubLink.Item.Name & " - " & mSubLink.Item.GetType.FullName)
                        'Next
                    End If
                Next currenLink
            Next currentGroup
        Next currentPage

    End Sub

    Public Sub OpenCloseBtn(ByVal BtnSv As Object, ByVal BtnEdt As Object, ByVal BtnDel As Object)
        Dim SqlStr As String
        SqlStr = "SELECT * FROM Users WHERE UserID LIKE '" & USERNAME & "%'"
        If Cn.State = ConnectionState.Closed Then Cn.Open()
        Dim DS As New DataSet
        Dim ADP20 As New OleDb.OleDbDataAdapter(SqlStr, Cn)
        DS.Clear()
        ADP20.Fill(DS, "TBL")
        ADP20.Dispose()
        BtnSv.Enabled = DS.Tables(0).Rows(0).Item("S1")
        BtnEdt.Enabled = DS.Tables(0).Rows(0).Item("S2")
        BtnDel.Enabled = DS.Tables(0).Rows(0).Item("S3")
        'BtnFind.Enabled = DS.Tables(0).Rows(0).Item("S4")
        'BtnPrnt.Enabled = DS.Tables(0).Rows(0).Item("S5")
        Cn.Close()
    End Sub

    Public Sub OpenCloseBtn1(ByVal FIELD As String)
        Dim SqlStr As String
        SqlStr = "SELECT " & FIELD & " FROM Users WHERE UserID LIKE '" & USERNAME & "%'"
        If Cn.State = ConnectionState.Closed Then Cn.Open()
        Dim DS As New DataSet
        Dim ADP20 As New OleDb.OleDbDataAdapter(SqlStr, Cn)
        DS.Clear()
        ADP20.Fill(DS, "TBL")
        ADP20.Dispose()
        '  BtnSv.Enabled = DS.Tables(0).Rows(0).Item(FIELD)
        TestOpen = DS.Tables(0).Rows(0).Item(FIELD)
        Cn.Close()
    End Sub

    Public Sub ClearObjects(ByVal frm As Form, ByVal gbx As Object)
        On Error Resume Next
        Dim txt As Control
        Dim chk As CheckBox
        For Each txt In gbx.Controls 'frm.Controls
            If gbx Is Nothing Then
                Exit Sub
            ElseIf TypeOf txt Is TextBox Or TypeOf txt Is ComboBox Then
                txt.Text = ""
            End If
        Next
    End Sub
    Public Function HANY(ByVal HANY1 As Decimal, ByVal HANY2 As String) As String
        On Error Resume Next
        Dim VPS As Decimal = 0
        Dim V As Decimal = 0
        Dim WORDINTEGER As String = ""
        Dim LE As String = ""
        Dim P As String = ""
        Dim PS As String = ""
        HANY = ""
        Dim POUNDS As String = ""
        Dim WORDPS As String = ""
        Dim DOLLAR As String = ""
        Dim SENT As String = ""
        Dim SENTS As String = ""
        Dim TON As String = ""
        Dim KG As String = ""
        Dim KGS As String = ""
        Select Case HANY2
            Case "Libya"
                LE = " دينـار "
                P = " درهم "
                PS = " دراهم "
                POUNDS = " دنانير "
                V = Int(Math.Abs(HANY1))
                VPS = Val(Right(Format(HANY1, "000000000000.00"), 2))
                WORDINTEGER = AmountWord(V)
                WORDPS = AmountWord(VPS)
                If WORDINTEGER <> "" And (VPS <= 2) Then HANY = WORDINTEGER & LE & " و " & WORDPS & P & "فقط لاغير "
                If WORDINTEGER <> "" And (VPS >= 3 And VPS <= 9) Then HANY = WORDINTEGER & LE & " و " & WORDPS & PS & "فقط لاغير "
                If WORDINTEGER <> "" And (VPS > 9) Then HANY = WORDINTEGER & LE & " و " & WORDPS & P & "فقط لاغير "
                If WORDINTEGER = "" And (VPS <= 2) Then HANY = WORDPS & P & "فقط لاغير "
                If WORDINTEGER = "" And (VPS >= 3 And VPS <= 9) Then HANY = WORDPS & PS & "فقط لاغير "
                If WORDINTEGER = "" And VPS > 9 Then HANY = WORDPS & P & "فقط لاغير "
                If WORDINTEGER = "" And VPS = 0 Then HANY = ""
                If WORDINTEGER <> "" And VPS = 0 Then HANY = WORDINTEGER & LE & "فقط لاغير "
            Case "USA"
                DOLLAR = " دولار "
                SENT = " سنتاً "
                SENTS = "سنتات"
                V = Int(System.Math.Abs(HANY1))
                VPS = Val(Right(Format(HANY1, "000000000000.00"), 2))
                WORDINTEGER = AmountWord(V)
                WORDPS = AmountWord(VPS)
                If WORDINTEGER <> "" And (VPS <= 2) Then HANY = WORDINTEGER & DOLLAR & " و " & WORDPS & SENT & "فقط لاغير "
                If WORDINTEGER <> "" And (VPS >= 3 And VPS <= 9) Then HANY = WORDINTEGER & DOLLAR & " و " & WORDPS & " " & SENTS & " " & "فقط لاغير "
                If WORDINTEGER <> "" And (VPS > 9) Then HANY = WORDINTEGER & DOLLAR & " و " & WORDPS & SENT & "فقط لاغير "
                If WORDINTEGER = "" And (VPS <= 2) Then HANY = WORDPS & SENT & "فقط لاغير "
                If WORDINTEGER = "" And (VPS >= 3 And VPS <= 9) Then HANY = WORDPS & " " & SENTS & " " & "فقط لاغير "
                If WORDINTEGER = "" And VPS > 9 Then HANY = WORDPS & SENT & "فقط لاغير "
                If WORDINTEGER = "" And VPS = 0 Then HANY = ""
                If WORDINTEGER <> "" And VPS = 0 Then HANY = WORDINTEGER & DOLLAR & "فقط لاغير "
            Case "WEIGHT"
                TON = " طن "
                KG = " كيلو جرام "
                KGS = "كيلو جرامات"
                V = Int(Math.Abs(HANY1))
                VPS = Val(Right(Format(HANY1, "000000000000.000"), 3))
                WORDINTEGER = AmountWord(V)
                WORDPS = AmountWord(VPS)
                If WORDINTEGER <> "" And (VPS <= 2) Then HANY = WORDINTEGER & TON & " و " & WORDPS & KG & "فقط لاغير "
                If WORDINTEGER <> "" And (VPS >= 3 And VPS <= 9) Then HANY = WORDINTEGER & TON & " و " & WORDPS & KGS & "فقط لاغير "
                If WORDINTEGER <> "" And (VPS > 9) Then HANY = WORDINTEGER & TON & " و " & WORDPS & KG & "فقط لاغير "
                If WORDINTEGER = "" And (VPS <= 2) Then HANY = WORDPS & KG & "فقط لاغير "
                If WORDINTEGER = "" And (VPS >= 3 And VPS <= 9) Then HANY = WORDPS & KGS & "فقط لاغير "
                If WORDINTEGER = "" And VPS > 9 Then HANY = WORDPS & KG & "فقط لاغير "
                If WORDINTEGER = "" And VPS = 0 Then HANY = ""
                If WORDINTEGER <> "" And VPS = 0 Then HANY = WORDINTEGER & TON & "فقط لاغير "
        End Select
    End Function
    Private Function AmountWord(ByVal AMOUNT As Decimal) As String
        On Error Resume Next
        Dim n As Decimal = 0
        Dim C1 As Decimal = 0
        Dim C2 As Decimal = 0
        Dim C3 As Decimal = 0
        Dim C4 As Decimal = 0
        Dim C5 As Decimal = 0
        Dim C6 As Decimal = 0
        Dim S1 As String = ""
        Dim S2 As String = ""
        Dim S3 As String = ""
        Dim S4 As String = ""
        Dim S5 As String = ""
        Dim S6 As String = ""
        Dim C As String = ""
        n = Int(AMOUNT)
        C = Format(n, "000000000000")
        C1 = Val(Mid(C, 12, 1))
        Select Case C1
            Case Is = 1 : S1 = "واحد"
            Case Is = 2 : S1 = "اثنان"
            Case Is = 3 : S1 = "ثلاثة"
            Case Is = 4 : S1 = "اربعة"
            Case Is = 5 : S1 = "خمسة"
            Case Is = 6 : S1 = "ستة"
            Case Is = 7 : S1 = "سبعة"
            Case Is = 8 : S1 = "ثمانية"
            Case Is = 9 : S1 = "تسعة"
        End Select

        C2 = Val(Mid(C, 11, 1))
        Select Case C2
            Case Is = 1 : S2 = "عشر"
            Case Is = 2 : S2 = "عشرون"
            Case Is = 3 : S2 = "ثلاثون"
            Case Is = 4 : S2 = "اربعون"
            Case Is = 5 : S2 = "خمسون"
            Case Is = 6 : S2 = "ستون"
            Case Is = 7 : S2 = "سبعون"
            Case Is = 8 : S2 = "ثمانون"
            Case Is = 9 : S2 = "تسعون"
        End Select

        If S1 <> "" And C2 > 1 Then S2 = S1 + " و" + S2
        If S2 = "" Then S2 = S1
        If C1 = 0 And C2 = 1 Then S2 = S2 + "ة"
        If C1 = 1 And C2 = 1 Then S2 = "احدى عشر"
        If C1 = 2 And C2 = 1 Then S2 = "اثنى عشر"
        If C1 > 2 And C2 = 1 Then S2 = S1 + " " + S2
        C3 = Val(Mid(C, 10, 1))
        Select Case C3
            Case Is = 1 : S3 = "مائة"
            Case Is = 2 : S3 = "مئتان"
            Case Is > 2 : S3 = Left(AmountWord(C3), Len(AmountWord(C3)) - 1) + "مائة"
        End Select
        If S3 <> "" And S2 <> "" Then S3 = S3 + " و" + S2
        If S3 = "" Then S3 = S2

        C4 = Val(Mid(C, 7, 3))
        Select Case C4
            Case Is = 1 : S4 = "الف"
            Case Is = 2 : S4 = "الفان"
            Case 3 To 10 : S4 = AmountWord(C4) + " آلاف"
            Case Is > 10 : S4 = AmountWord(C4) + " الف"
        End Select
        If S4 <> "" And S3 <> "" Then S4 = S4 + " و" + S3
        If S4 = "" Then S4 = S3
        C5 = Val(Mid(C, 4, 3))
        Select Case C5
            Case Is = 1 : S5 = "مليون"
            Case Is = 2 : S5 = "مليونان"
            Case 3 To 10 : S5 = AmountWord(C5) + " ملايين"
            Case Is > 10 : S5 = AmountWord(C5) + " مليون"
        End Select
        If S5 <> "" And S4 <> "" Then S5 = S5 + " و" + S4
        If S5 = "" Then S5 = S4

        C6 = Val(Mid(C, 1, 3))

        Select Case C6
            Case Is = 1 : S6 = "مليار"
            Case Is = 2 : S6 = "ملياران"
            Case Is > 2 : S6 = AmountWord(C6) + " مليار"
        End Select
        If S6 <> "" And S5 <> "" Then S6 = S6 + " و" + S5
        If S6 = "" Then S6 = S5
        AmountWord = S6
    End Function

    Public Function MydateWord(ByVal MYDATE As Date) As String
        On Error Resume Next
        Dim n As Integer = 0
        Dim C1 As Decimal = 0
        Dim C2 As Decimal = 0
        Dim C3 As Decimal = 0
        Dim S1 As String = ""
        Dim S2 As String = ""
        Dim C As String = ""
        MydateWord = ""
        C = Format(MYDATE, "dd-mm-yyyy")
        C1 = MYDATE.Day
        C2 = MYDATE.Month
        C3 = MYDATE.Year
        Select Case C1
            Case Is = 1 : S1 = "الاول"
            Case Is = 2 : S1 = "الثانى"
            Case Is = 3 : S1 = "الثالث"
            Case Is = 4 : S1 = "الرابع"
            Case Is = 5 : S1 = "الخامس"
            Case Is = 6 : S1 = "السادس"
            Case Is = 7 : S1 = "السابع"
            Case Is = 8 : S1 = "الثامن"
            Case Is = 9 : S1 = "التاسع"
            Case Is = 10 : S1 = "العاشر"
            Case Is = 11 : S1 = "الحادى عشر"
            Case Is = 12 : S1 = "الثانى عشر"
            Case Is = 13 : S1 = "الثالث عشر"
            Case Is = 14 : S1 = "الرابع عشر"
            Case Is = 15 : S1 = "الخامس عشر"
            Case Is = 16 : S1 = "السادس عشر"
            Case Is = 17 : S1 = "السابع عشر"
            Case Is = 18 : S1 = "الثامن عشر"
            Case Is = 19 : S1 = "التاسع عشر"
            Case Is = 20 : S1 = "العشرون"
            Case Is = 21 : S1 = "الواحد والعشرون"
            Case Is = 22 : S1 = " الثانى والعشرون"
            Case Is = 23 : S1 = "الثالث والعشرون"
            Case Is = 24 : S1 = " الرابع والعشرون"
            Case Is = 25 : S1 = " الخامس والعشرون"
            Case Is = 26 : S1 = "السادس والعشرون"
            Case Is = 27 : S1 = "السابع والعشرون"
            Case Is = 28 : S1 = "الثامن والعشرون"
            Case Is = 29 : S1 = "التاسع والعشرون"
            Case Is = 30 : S1 = "الثلاثون"
            Case Is = 31 : S1 = "الواحد والثلاثون"
        End Select

        Select Case C2
            Case Is = 1 : S2 = "يناير"
            Case Is = 2 : S2 = "فبراير"
            Case Is = 3 : S2 = "مارس"
            Case Is = 4 : S2 = "ابريل"
            Case Is = 5 : S2 = "مايو"
            Case Is = 6 : S2 = "يونية"
            Case Is = 7 : S2 = "يوليو"
            Case Is = 8 : S2 = "اغسطس"
            Case Is = 9 : S2 = "سبتمبر"
            Case Is = 10 : S2 = "اكتوبر"
            Case Is = 11 : S2 = "نوفمبر"
            Case Is = 12 : S2 = "ديسمبر"
        End Select
        MydateWord = Format(MYDATE, "dddd") & " الموافق " & S1 & " من  شهر " & S2 & " سنة " & AmountWord(CDec(C3)) & " ميلادية "
    End Function
    Public Function MytimeWord(ByVal MYTIME As DateTime) As String
        On Error Resume Next
        Dim n As Integer = 0
        Dim C1 As Decimal = 0
        Dim C2 As Decimal = 0
        Dim C3 As Decimal = 0
        Dim C4 As String = ""
        Dim S1 As String = ""
        Dim S2 As String = ""
        Dim S3 As String = ""
        Dim S4 As String = ""
        Dim S5 As String = ""
        Dim C As DateTime
        MytimeWord = ""
        C = Format(MYTIME, "hh:mm:ss tt")
        C1 = Format(C, "ss")
        C2 = Format(C, "mm")
        C3 = Format(C, "hh")
        C4 = Format(C, "tt")
        Select Case C4
            Case "ص" : S4 = "صباحاَ"
            Case "م" : S4 = "مساءً"
            Case "AM" : S4 = "صباحاَ"
            Case "PM" : S4 = "مساءً"
        End Select
        Select Case C1
            Case Is = 0 : S3 = ""
            Case Is = 1 : S3 = " ثانية واحدة "
            Case Is = 2 : S3 = " ثانيتان"
            Case 3 To 10 : S3 = AmountWord(C1) + " ثوان"
            Case Is > 10 : S3 = AmountWord(C1) + " ثانية"
        End Select
        Select Case C2
            Case Is = 0 : S1 = ""
            Case Is = 1 : S1 = " دقيقة واحدة "
            Case Is = 2 : S1 = " دقيقتان "
            Case 3 To 10 : S1 = AmountWord(C2) + " دقائق "
            Case Is > 10 : S1 = AmountWord(C2) + " دقيقة "
        End Select
        If S1 <> "" And S3 <> "" Then S5 = S1 + " و" + S3
        If S1 = "" Then S5 = S3
        Select Case C3
            Case Is = 0 : S2 = ""
            Case Is = 1 : S2 = "الواحدة"
            Case Is = 2 : S2 = "الثانية"
            Case Is = 3 : S2 = "الثالثة"
            Case Is = 4 : S2 = "الرابعة"
            Case Is = 5 : S2 = "الخامسة"
            Case Is = 6 : S2 = "السادسة"
            Case Is = 7 : S2 = "السابعة"
            Case Is = 8 : S2 = "الثامنة"
            Case Is = 9 : S2 = "التاسعة"
            Case Is = 10 : S2 = "العاشرة"
            Case Is = 11 : S2 = "الحادية عشر"
            Case Is = 12 : S2 = "الثانية عشر"
        End Select
        If S2 <> "" And S1 <> "" And S3 <> "" Then S5 = S2 + " و" + S1 + " و" + S3
        If S2 <> "" And S1 <> "" And S3 = "" Then S5 = S2 + " و" + S1
        If S2 <> "" And S1 = "" And S3 <> "" Then S5 = S2 + " و" + S3
        If S2 <> "" And S1 = "" And S3 = "" Then S5 = S2
        MytimeWord = " الساعة " & S5 & " " & S4
    End Function
End Module
