Public Class Form1
  Private aeDesktop As AutomationElement
    Private aeCalculator As AutomationElement

    Private ae5Btn As AutomationElement
    Private aeAddBtn As AutomationElement
    Private aeEqualsBtn As AutomationElement

    Private strDesktopPath As String = "C:\Users\" & Environ("USERNAME") & "\Desktop\"
    Private p As Process
  Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim ms As Int32 = 100
        Dim desktop As UIAutomationHelper = New UIAutomationHelper
        ''Launch Windows Calculator (calc.exe) and find the main window
        Dim window As AutomationElement = desktop.LaunchApplication("Calc", "", "Calculator")
        ''Get a reference to the Edit TextBox (ControlId = "96")
        Dim textbox As AutomationElement = desktop.GetElementById(window, "150")
        ''Get a reference to the RadioButton "Dec" (AutomationId = "314")
        Dim radioButtonDecimal As AutomationElement = desktop.GetElementById(window, "314")

        ''If this RadioButton can not be found, the calculator uses the Standard view
        ''Call the Scientfic menu to make more options visible

        If IsNothing(radioButtonDecimal) Then
            '' Need to expand the menu
            
            '' New List(Of String) From {"View", "Scientific"}

            ''desktop.OpenMenu(window, "View")
           
            ''desktop.ExecuteMenuItem(window, "Scientific")
            desktop.TraverseMenu(window, New List(Of String) From {"View", "Scientific"})
        End If

        ''Focus the Edit TextBox and enter the text 54 

        Thread.Sleep(ms)

        ''Press button "3" (ControlId = "85") so the value will be 133

        desktop.InvokeElementById(window, "133")
        Thread.Sleep(ms)
        ''Press the "+" button (ControlId = "5D") = 93

        desktop.InvokeElementByName(window, "Add")
        Thread.Sleep(ms)

        ''Enter the second value. Press button "6" (ControlId = "88")= 136
        desktop.InvokeElementByName(window, "6")
        Thread.Sleep(ms)
        ''Enter the second value. Press button "7" (ControlId = "131")
        desktop.InvokeElementByName(window, "7")
        Thread.Sleep(ms)
        ''Press the "=" button to calculate the result (ControlId = "79")=121
        desktop.InvokeElementById(window, "121")
        Thread.Sleep(ms)
        '' USe the TraverseMenu
        desktop.TraverseMenu(window, New List(Of String) From {"Edit", "Copy"})
        'desktop.OpenMenu(window, "Edit")
        'Thread.Sleep(ms)
        'desktop.ExecuteMenuItem(window, "Copy")
        ''Press F8 to change the numerical system from decimal to binary
        ''SendKeys.SendWait("{F8}")

        Me.txtBoxResult1.Text = "3+67 = " + System.Windows.Forms.Clipboard.GetText()

        If Not IsNothing(window) Then desktop.CloseWindow(window)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ''Launch Windows Calculator (calc.exe) and find the main window
        Dim desktop As UIAutomationHelper = New UIAutomationHelper()
        Dim window As AutomationElement = desktop.LaunchApplication("Calc", "", "Calculator")

        desktop.DumpControlNameId(window)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim UIAhelper As UIAutomationHelper = New UIAutomationHelper()
        Dim desktop As AutomationElement = UIAhelper.GetDesktop
        If IsNothing(desktop) Then
            Exit Sub
        End If

        Try

            Dim tab As AutomationElement = UIAhelper.GetElementContainsName(desktop, txtWName.Text)
            If Not IsNothing(tab) Then


                UIAhelper.DumpControlNameId(tab)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim UIAhelper As UIAutomationHelper = New UIAutomationHelper()
        Dim desktop As AutomationElement = UIAhelper.GetDesktop
        If IsNothing(desktop) Then
            Exit Sub
        End If

        Try

            Dim element As AutomationElement = UIAhelper.GetElementById(desktop, txtElementID.Text)
            If Not IsNothing(element) Then


                UIAhelper.DumpControlNameId(element)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
  End Class
