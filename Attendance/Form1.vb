

Imports OpenQA.Selenium

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim table As New List(Of Tuple(Of DateTime, DateTime))
        Dim signin, signout As DateTime
        Dim options As New Chrome.ChromeOptions()
        Dim weekend As Boolean = False
        Dim TotalDuration As New TimeSpan
        Dim DayDuration As New TimeSpan
        Dim TotalReduction As New Integer


        options.AddArguments("--no-sandbox")

        Dim MyBrowser As New Chrome.ChromeDriver("D:\", options)
        MyBrowser.Url = "http://mcssrv-hrn/HR/Login.aspx"

        MyBrowser.Navigate()


        MyBrowser.FindElementById("Txt_UserName").SendKeys("s.abood")
        MyBrowser.FindElementById("Txt_Password").SendKeys("P@$$w0rd!")

        MyBrowser.FindElementById("Btn_Login").Click()
        MyBrowser.Url = "http://mcssrv-hrn/HR/Pages/Timesheet/TimesheetDetails/EmployeeTimesheetDetails.aspx"
        MyBrowser.Navigate()
        MyBrowser.FindElementById("ContentPlaceHolder1_DDL_Year").SendKeys(DateTime.Now.Year)
        MyBrowser.FindElementById("ContentPlaceHolder1_DDL_Month").SendKeys(DateTime.Now.ToString("MMM"))
        MyBrowser.FindElementById("ContentPlaceHolder1_Btn_Refresh").Click()

        For i = 0 To DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month - 1) - 1

            If MyBrowser.FindElementById("ContentPlaceHolder1_GV_TimesheetDetails_Pnl_AttendanceDay_" + i.ToString).GetCssValue("background-color") = "rgba(211, 211, 211, 1)" OrElse String.IsNullOrEmpty(MyBrowser.FindElementById("ContentPlaceHolder1_GV_TimesheetDetails_Lbl_SignIn_" + i.ToString).Text) OrElse String.IsNullOrEmpty(MyBrowser.FindElementById("ContentPlaceHolder1_GV_TimesheetDetails_Lbl_SignIn_" + i.ToString).Text) Then
                Continue For
            End If

            signin = MyBrowser.FindElementById("ContentPlaceHolder1_GV_TimesheetDetails_Lbl_SignIn_" + i.ToString).Text
            signout = MyBrowser.FindElementById("ContentPlaceHolder1_GV_TimesheetDetails_Lbl_SignOut_" + i.ToString).Text
            DayDuration = signout.Subtract(signin)
            TotalReduction += 8 - DayDuration.Hours
            TotalDuration += DayDuration

            Console.WriteLine(TotalDuration)
            table.Add(New Tuple(Of DateTime, DateTime)(signin, signout))


        Next





    End Sub
End Class
