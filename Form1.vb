Imports System.Data.SqlClient
Imports System.IO
Imports System.ServiceProcess



Public Class Ocean
    Public MyTaskName As String = "OceanTask"
    Public seldrive As String = "D"
    Public foloc As String
    Public Backup As String
    Public sqlfile As String
    Public namefile As String

    Public backbat As String = seldrive & ":\baktools\backup.bat"

    Public remotedailytask As String
    Public Sub New()
        InitializeComponent()
        cbx.SelectedIndex = 0
        chb2.Checked = True
        fCB.Checked = True
        Dim controller As New ServiceController("MSSQLSERVER")

        If controller.Status = ServiceControllerStatus.Running Then sst.Text = "Running"
        If controller.Status = ServiceControllerStatus.Stopped Then sst.Text = "Stopped"
        If controller.Status = ServiceControllerStatus.Stopped Then sst.ForeColor = Color.Red
        If controller.Status = ServiceControllerStatus.Running Then sst.ForeColor = Color.Honeydew
        companyname.Text = Environment.MachineName

        ' DateTimePicker1.ResetText = "17:00"

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Call Maintask()


    End Sub


    Public Sub Maintask()
        Dim sqlopencmd As String
        MyTaskName = "OceanTask"
        seldrive = cbx.Text
        If cbx.Text = "Auto" Then seldrive = "D"
        sqlopencmd = "Data Source=.;Initial Catalog=master;Persist Security Info=True;Integrated Security=SSPI"
        foloc = seldrive & ":\baktools"
        Backup = seldrive & ":\Backup"
        sqlfile = seldrive & ":\baktools\backup.sql"
        namefile = seldrive & ":\baktools\name.bat"
        backbat = seldrive & ":\baktools\backup.bat"
        Dim runinSM As String = seldrive & ":\baktools\runinSM.bat"
        remotedailytask = foloc & "\RunToCreateTask.bat"
        Dim constr As String = New SqlConnection(sqlopencmd).ConnectionString





        On Error GoTo err1
        If (Not System.IO.Directory.Exists(foloc)) Then
            System.IO.Directory.CreateDirectory(foloc)
        End If

        System.IO.File.WriteAllBytes(foloc & " \rar.exe", My.Resources.Rar)
        System.IO.File.WriteAllBytes(foloc & "\PKUNZIP.exe", My.Resources.PKUNZIP)
        System.IO.File.WriteAllBytes(foloc & "\PKZIP.exe", My.Resources.PKZIP)
        System.IO.File.WriteAllText(foloc & "\silentmode.vbs", My.Resources.silentmode)
        System.IO.File.WriteAllText(foloc & "\sm.bat", My.Resources.sm)



        Dim folstt As String
        If (System.IO.File.Exists(backbat)) And (System.IO.File.Exists(backbat)) And (System.IO.File.Exists(backbat)) And (System.IO.Directory.Exists(foloc)) Then
            folstt = "updated"
        Else
            folstt = "created"
        End If
        Dim SqlConSt As String = "SQL Server services Is Not enabled Or Not found  "
        ' Dim varCollection As  = Nothing
        'Dim delimiter As String = varCollection("user::Delimiter").value.tostring()
        On Error GoTo err1
        Dim sqlcode As String = String.Empty
        Dim Ntxt As String = String.Empty
        Dim INtxt As String = String.Empty
        Dim btxt As String = String.Empty
        Dim Driva As String = seldrive
        sqlcode = "declare @name  nvarchar(50),@Fname  nvarchar(50),@Tname nvarchar(50)
declare @date nvarchar(11) 
declare @dd nvarchar(10) 
declare @MM nvarchar(10) 
declare @yy nvarchar(10) 
set @date=getdate() 
set @dd= day(@date) 
set @mm=month(@date) 
set @yy=year(@date) 
Declare cur Cursor 
for select name from sys.databases where name not in ('master','ODBOLD','tempdb','model','msdb','master', 'tempdb', 'model', 'msdb','reportserver','reportservertempdb') 
open cur
fetch next from cur into @name 
While @@FETCH_STATUS = 0
begin
	set @Fname ='" & Driva & ":\Baktools\'+@name+'_'+ @dd +'_'+ @mm +'_'+ @yy+'.Bak' 
	set @Tname= @name+'-Full Database Backup'
	BACKUP DATABASE @name TO  DISK = @Fname WITH NOFORMAT, NOINIT,  NAME = @Tname, SKIP, NOREWIND, NOUNLOAD,  STATS = 10
	fetch next from cur into @name
end
CLOSE cur
DEALLOCATE cur"

        Ntxt = "rem @echo off

set thedate=%date%
set yyyy=%thedate:~10,4%
set dd=%thedate:~7,2%
set mm=%thedate:~4,2%
set hh=%time:~0,2%
move /y " & Driva & ":\BakTools\log\backlog.log " & Driva & ":\Backup\log\bcklog_%dd%.log
move /y " & Driva & ":\Baktools\*.bak " & Driva & ":\Backup\
rem " & Driva & ":\BaKTools\rar a -df -ag_dd_mm_yyyy_hh_mm " & Driva & ":\backup\Bakup.rar " & Driva & ":\Backup\*.bak

" & Driva & ":\BaKTools\rar a -df  " & Driva & ":\backup\Bakup.rar " & Driva & ":\Backup\*.bak
move /y " & Driva & ":\Backup\Bakup.rar " & Driva & ":\Backup\Bakup_%dd%_%mm%_%yyyy%_" & companyname.Text & ".rar
xcopy " & Driva & ":\Backup\Bakup_%dd%_%mm%_%yyyy%_" & companyname.Text & ".rar " & Driva & ":\upload\ /y /s /d"






        Using sw As StreamWriter = New StreamWriter(sqlfile)
            sw.WriteLine(sqlcode)
        End Using
        Using sw As StreamWriter = New StreamWriter(namefile)
            sw.WriteLine(Ntxt)
        End Using
        Using sw As StreamWriter = New StreamWriter(backbat)

            btxt = "
if not exist """ & Driva & "\backup"" mkdir " & Driva & ":\backup
if not exist """ & Driva & "\Baktools\Log"" mkdir " & Driva & ":\Baktools\Log
if not exist """ & Driva & "\Backup\Log"" mkdir " & Driva & ":\backup\Log
if not exist """ & Driva & "\Upload"" mkdir " & Driva & ":\Upload
SQLCMD.EXE -i " & Driva & ":\Baktools\backup.sql -o " & Driva & ":\Baktools\log\backlog.log

call " & Driva & ":\Baktools\name.bat
call " & Driva & ":\Baktools\CopyBack.bat"


            sw.WriteLine(btxt)
        End Using


        'err2:
        '  MsgBox("SQL Server services is not enabled or not found  " & SqlConSt)
err1:
        Dim taskcreated As String = " and a daily task assigned to run daily every " & DateTimePicker1.Text

        If fCB.Checked = True Then
            Call Taskcreatornomsgbox()

        End If



        If (System.IO.File.Exists(backbat)) And (System.IO.File.Exists(backbat)) And (System.IO.File.Exists(backbat)) And (System.IO.Directory.Exists(foloc)) Then
            MsgBox("Baktools folder " & folstt & taskcreated & " successfully! ", MsgBoxStyle.Information, Title:="Done")
        Else
            MsgBox("Faild to create Baktools folder! " & SqlConSt, MsgBoxStyle.Exclamation, Title:="Error")
        End If

        If chb2.Checked = True Then Call Taskrunnernomsg()


    End Sub


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "HH:mm"
        DateTimePicker1.ShowUpDown = True
        DateTimePicker1.Value = "1/1/2020 18:00:00"
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub
    Public Sub Taskcreator()
        If SMCB.Checked = False Then backbat = seldrive & ":\baktools\backup.bat" Else backbat = seldrive & ":\baktools\sm.bat"
        If (System.IO.File.Exists(backbat)) Then
            Using tService As New TaskService()
                Dim mytaskstt As String

                Dim tTask As Task = tService.GetTask(MyTaskName)
                If tTask Is Nothing Then mytaskstt = "created " Else mytaskstt = "updated"
                Dim tDefinition As TaskDefinition = tService.NewTask
                tDefinition.RegistrationInfo.Description =
                   "This task was created to backup all databases on this sql server  created by OceanSoft Team"

                Dim tTrigger As New DailyTrigger()

                tTrigger.StartBoundary = "1899-12-30T" & DateTimePicker1.Text & ":00"

                tDefinition.Triggers.Add(tTrigger)

                tDefinition.Actions.Add(New ExecAction(backbat))


                tService.RootFolder.RegisterTaskDefinition(MyTaskName,
                   tDefinition)
                tTask = tService.GetTask(MyTaskName)
                'Dim tTask As Task = tService.GetTask(MyTaskName)
                If tTask Is Nothing Then
                    MsgBox("Faild to create task!")

                Else
                    MsgBox("Task " & mytaskstt & " Successfully")

                End If
            End Using
        Else
            MsgBox("Faild to create task!, Please create the Baktools folder first")
        End If
    End Sub
    Public Sub Taskrunner()
        If (System.IO.File.Exists(backbat)) Then
            Using tService As New TaskService()
                Dim tTask As Task = tService.GetTask(MyTaskName)
                If tTask Is Nothing Then
                    Select Case MsgBox("The Task is not found!" & vbNewLine & " Would you like to create A Daily task every " & DateTimePicker1.Text & " ?", MsgBoxStyle.YesNo, "Tasks")
                        Case MsgBoxResult.Yes
                            Call Taskcreator()
                            Call Taskrunner()

                    End Select
                    '  MsgBox("The Task is not found , Would you like to create A Daily task every " & DateTimePicker1.Text, MsgBoxStyle.YesNo)
                    ' If MsgBoxResult.Yes Then Call Taskcreator() Else MsgBox("it will be not")
                    'If vbYesNo = True Then MsgBox("it will create") 'Call Taskcreator()
                Else
                    Shell("schtasks /Run  /TN " & MyTaskName)
                    Dim notif As New NotifyIcon
                    notif.Visible = True
                    notif.ShowBalloonTip(3000, "Backup Task Sarted", "Please wait untill the backup task finish", ToolTipIcon.Info)

                    Return

                End If
            End Using
            'End Using
        Else
            MsgBox("Create Baktools and a Daily Task first")
        End If
    End Sub




    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call Taskcreator()
        '   MsgBox("Task Created Successfully")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call Taskrunner()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Shell("control schedtasks")
    End Sub

    Public Sub Taskcreatornomsgbox()
        If SMCB.Checked = False Then backbat = seldrive & ":\baktools\backup.bat" Else backbat = seldrive & ":\baktools\sm.bat"

        If (System.IO.File.Exists(backbat)) Then
            Using tService As New TaskService()
                Dim mytaskstt As String

                Dim tTask As Task = tService.GetTask(MyTaskName)
                If tTask Is Nothing Then mytaskstt = "created " Else mytaskstt = "updated"
                Dim tDefinition As TaskDefinition = tService.NewTask
                tDefinition.RegistrationInfo.Description =
                   "This task was created to backup all databases on this sql server  created by OceanSoft Team"

                Dim tTrigger As New DailyTrigger()

                tTrigger.StartBoundary = "1899-12-30T" & DateTimePicker1.Text & ":00"

                tDefinition.Triggers.Add(tTrigger)

                tDefinition.Actions.Add(New ExecAction(backbat))


                tService.RootFolder.RegisterTaskDefinition(MyTaskName,
                   tDefinition)
                tTask = tService.GetTask(MyTaskName)
                'Dim tTask As Task = tService.GetTask(MyTaskName)
                If tTask Is Nothing Then
                    ' MsgBox("Faild to create task!")

                Else
                    ' MsgBox("Task " & mytaskstt & " Successfully")

                End If
            End Using
        Else

        End If
    End Sub
    Public Sub Taskrunnernomsg()
        If SMCB.Checked = False Then backbat = seldrive & ":\baktools\backup.bat" Else backbat = seldrive & ":\baktools\sm.bat"

        If (System.IO.File.Exists(backbat)) Then
            Using tService As New TaskService()
                Dim tTask As Task = tService.GetTask(MyTaskName)
                If tTask Is Nothing Then
                    Select Case MsgBox("The Task is not found , Would you like to create A Daily task every " & DateTimePicker1.Text & " ?", MsgBoxStyle.YesNo, "Tasks")
                        Case MsgBoxResult.Yes
                            Call Taskcreator()

                    End Select
                    '  MsgBox("The Task is not found , Would you like to create A Daily task every " & DateTimePicker1.Text, MsgBoxStyle.YesNo)
                    ' If MsgBoxResult.Yes Then Call Taskcreator() Else MsgBox("it will be not")
                    'If vbYesNo = True Then MsgBox("it will create") 'Call Taskcreator()
                Else Shell("schtasks /Run  /TN " & MyTaskName)
                    Return

                End If
            End Using
            'End Using
        Else
            '  MsgBox("Create Baktools and a Daily Task first")
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        sst.Text = ""
        Dim controller As New ServiceController("MSSQLSERVER")
        If controller.Status = ServiceControllerStatus.Stopped Then
            controller.Start()
            System.Threading.Thread.Sleep(5000)
            If controller.Status = ServiceControllerStatus.Stopped Then sst.ForeColor = Color.Red
            If controller.Status = ServiceControllerStatus.Running Then sst.ForeColor = Color.Honeydew
            sst.Text = "Task started"
            If sst.Text = "Task started" Then sst.ForeColor = Color.Honeydew
        Else
            sst.Text = "Already Running"
        End If

    End Sub



    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        sst.Text = ""
        Dim controller As New ServiceController("MSSQLSERVER")
        If controller.Status = ServiceControllerStatus.Running Then
            controller.Stop()
            If controller.Status = ServiceControllerStatus.Stopped Then sst.ForeColor = Color.Red
            If controller.Status = ServiceControllerStatus.Running Then sst.ForeColor = Color.Honeydew
            System.Threading.Thread.Sleep(5000)
            sst.ForeColor = Color.Red
            sst.Text = "Stopped"


        Else
            sst.Text = "Already stopped"

        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        sst.Text = ""
        Dim controller As New ServiceController("MSSQLSERVER")
        If controller.Status = ServiceControllerStatus.Running Then controller.Stop()
        If controller.Status = ServiceControllerStatus.Stopped Then controller.Start()

        sst.Text = "Resterted"
        If controller.Status = ServiceControllerStatus.Stopped Then sst.ForeColor = Color.Red
        If controller.Status = ServiceControllerStatus.Running Then sst.ForeColor = Color.Honeydew
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        Dim sqlopencmd As String = "Data Source=.;Initial Catalog=master;Persist Security Info=True;Integrated Security=SSPI"

        Dim constr As String = New SqlConnection(sqlopencmd).ConnectionString
        Using con As New SqlConnection(constr)
            Using cmd As New SqlCommand("SELECT command, percent_complete,total_elapsed_time, estimated_completion_time, start_time
  FROM sys.dm_exec_requests
  WHERE command  ='BACKUP DATABASE'")

            End Using

        End Using

        Dim table As New DataTable()

        Dim connectionString As String = "Server=.;Trusted_Connection=True;
"
        Dim constring As String = "SELECT command, percent_complete,total_elapsed_time, estimated_completion_time, start_time
  FROM sys.dm_exec_requests
  WHERE command  ='BACKUP DATABASE'"
        Using con = New SqlConnection(connectionString)
            con.Open()
            Dim dsMyName As New DataSet

            Using command = New SqlCommand(constring, con)
                Using da = New SqlDataAdapter(command)
                    '  MONI.Text = (command.ToString)
                    da.SelectCommand = command
                    Try
                        da.Fill(dsMyName)
                        Dim progressvalueas As String = dsMyName.Tables(0).Rows(0).Item("percent_complete").ToString
                        '  progressvalueas = progressvalueas.to
                        progressvalueas = CInt(progressvalueas)
                        If progressvalueas = "System.Data.SqlClient.SqlCommand" Then progressvalueas = ""
                        'MONI.Text = progressvalueas
                    Catch
                    End Try
                End Using
            End Using
        End Using


    End Sub

    Private Sub sst_TextChanged(sender As Object, e As EventArgs) Handles sst.TextChanged

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class

