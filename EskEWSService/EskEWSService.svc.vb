' NOTE: You can use the "Rename" command on the context menu to change the class name "EskEWSService" in code, svc and config file together.
' NOTE: In order to launch WCF Test Client for testing this service, please select EskEWSService.svc or EskEWSService.svc.vb at the Solution Explorer and start debugging.
Imports System.Data.SqlClient
Imports System.IdentityModel.Diagnostics
Imports Microsoft.Exchange.WebServices.Data

Public Class EskEWSService
    Implements IEskEWSService
    Dim connString As String = "Data Source=.;server=sqlcls;Initial Catalog=ScaDataDB;User ID=sa;Password=sqladmin"

    Public Sub New()
    End Sub

    Private Function IsAuthorised(MyLogin As String, MyService As String) As Boolean
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка - есть ли у данного пользователя право на использование сервиса
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim ds As New DataSet()
        Dim MyAuth As Boolean = False

        Try
            MySQLStr = "dbo.spp_Services_GetAuthInfo"
            Using MyConn As SqlConnection = New SqlConnection(connString)
                Try
                    Using cmd As SqlCommand = New SqlCommand(MySQLStr, MyConn)
                        cmd.CommandType = CommandType.StoredProcedure
                        cmd.CommandTimeout = 1800
                        cmd.Parameters.AddWithValue("@MyLogin", MyLogin)
                        cmd.Parameters.AddWithValue("@MyService", MyService)
                        Using da As New SqlDataAdapter()
                            da.SelectCommand = cmd
                            da.Fill(ds)
                            If ds.Tables(0).Rows.Count <> 0 Then
                                MyAuth = True
                            End If
                        End Using
                    End Using
                Catch ex As Exception
                    EventLog.WriteEntry("ESKEWSServices", "IsAuthorised --1--> " & ex.Message)
                Finally
                    MyConn.Close()
                End Try
            End Using
        Catch ex As Exception
            EventLog.WriteEntry("ESKEWSServices", "IsAuthorised --2--> " & ex.Message)
        End Try
        Return MyAuth
    End Function


    Public Async Function CreateCalendarEventAsync(ByVal MyEvent As CreateCalendarEventType) As Threading.Tasks.Task(Of String) Implements IEskEWSService.CreateCalendarEventAsync
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание события в календаре
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MyLogin As String
        Dim MyService As String
        Dim MyId As String

        MyLogin = MyEvent.Login
        MyService = "EskEWSCreateCalendarEvent"

        If IsAuthorised(MyLogin, MyService) Then
            '------------в случае авторизации - создание события
            Try
                MyId = Await FCreateCalendarEvent(MyEvent)
                Return MyId
            Catch ex As Exception
                Return ""
            End Try
        Else
            Return ""
        End If
    End Function

    Private Async Function FCreateCalendarEvent(ByVal MyEvent As CreateCalendarEventType) As Threading.Tasks.Task(Of String)
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Функция создания события в календаре
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MySrvc As Microsoft.Exchange.WebServices.Data.ExchangeService
        Dim MyId As String

        Try
            MySrvc = GetService("Backup.Operator@elektroskandia.ru", "B@ckUp0per")
            If IsNothing(MySrvc) Then
                Return ""
            End If
            If MyEvent.CalendarEventIDOld.Equals("") = False Then
                '-----Удаление старой записи с ID
                Try
                    MyId = Await FDeleteCalendarEvent(MyEvent.CalendarEventIDOld)
                Catch ex As Exception
                End Try
            End If
            Dim appointment As Appointment = New Appointment(MySrvc)
            appointment.Subject = MyEvent.Subject
            appointment.Body = MyEvent.Body
            appointment.Start = MyEvent.Start
            appointment.End = MyEvent.Finish
            appointment.ReminderMinutesBeforeStart = 30
            appointment.Save(New FolderId(WellKnownFolderName.Calendar, New Mailbox(MyEvent.Email)))
            Dim Item As Item = Item.Bind(MySrvc, appointment.Id, New PropertySet(ItemSchema.Subject))
            Return appointment.Id.ToString
        Catch ex As Exception
            Return ""
        End Try
        Return ""
    End Function

    Friend Shared Function GetService(emailAddress As String, userPass As String) As Microsoft.Exchange.WebServices.Data.ExchangeService
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Функция получения сервиса
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim Service As Microsoft.Exchange.WebServices.Data.ExchangeService = Nothing
        Try
            Service = New Microsoft.Exchange.WebServices.Data.ExchangeService(Microsoft.Exchange.WebServices.Data.ExchangeVersion.Exchange2007_SP1)
            Service.Credentials = New WebCredentials(emailAddress, userPass)
            Service.AutodiscoverUrl(emailAddress, AddressOf MyCallBack)
        Catch ex As Microsoft.Exchange.WebServices.Data.AutodiscoverLocalException
        End Try
        Return Service
    End Function

    Public Shared Function MyCallBack(ByVal url As String) As Boolean
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// CallBack Функции получения сервиса
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////

        Return True
    End Function



    Public Async Function DeleteCalendarEventAsync(ByVal MyEvent As DeleteCalendarEventType) As Threading.Tasks.Task(Of String) Implements IEskEWSService.DeleteCalendarEventAsync
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление события в календаре
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MyLogin As String
        Dim MyService As String
        Dim MyId As String

        MyLogin = MyEvent.Login
        MyService = "EskEWSCreateCalendarEvent"

        If IsAuthorised(MyLogin, MyService) Then
            '------------в случае авторизации - удаление события
            Try
                MyId = Await FDeleteCalendarEvent(MyEvent.CalendarEventIDOld)
                Return MyId
            Catch ex As Exception
                Return ""
            End Try
        Else
            Return ""
        End If
    End Function

    Private Async Function FDeleteCalendarEvent(ByVal MyID As String) As Threading.Tasks.Task(Of String)
        '//////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Функция удаления события в календаре
        '//
        '//////////////////////////////////////////////////////////////////////////////////////////
        Dim MySrvc As Microsoft.Exchange.WebServices.Data.ExchangeService = Nothing

        If MyID.Equals("") = False Then
            Try
                MySrvc = GetService("Backup.Operator@elektroskandia.ru", "B@ckUp0per")
                If IsNothing(MySrvc) Then
                    Return ""
                End If
                Dim appointmentId As String
                appointmentId = MyID
                Dim appointment As Appointment = Appointment.Bind(MySrvc, appointmentId, New PropertySet())
                appointment.Delete(DeleteMode.MoveToDeletedItems)
                Return "Success"
            Catch ex As Exception
                Return ""
            End Try
            Return ""
        End If
        Return "Success"
    End Function
End Class
