Imports System.Windows.Automation
Imports System.Threading

Public Class UIAutomationHelper
    Shared Sub New()

    End Sub

    

    Public Function LaunchApplication(ByVal filename As String, arguments As String, Optional windowtitle As String = "") As AutomationElement
        Dim calcAutomationElement As AutomationElement
        Dim ct As Integer
        Dim process As Process = New Process()
        Try
            process.StartInfo.FileName = filename
            process.StartInfo.Arguments = arguments
            process.Start()


            ''Loop and wait for process to level out
            Do
                calcAutomationElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, New PropertyCondition(AutomationElement.NameProperty, windowtitle))
                ct += 1
                Thread.Sleep(100)
            Loop While IsNothing(calcAutomationElement) And ct < 50
            If IsNothing(calcAutomationElement) Then Return Nothing
            Return calcAutomationElement ''AutomationElement.FromHandle(process.MainWindowHandle)
        Catch e As Exception
            Return Nothing
            'MsgBox(e.Message)
        End Try
    End Function
    Public Function GetDesktop() As AutomationElement
        Return AutomationElement.RootElement
    End Function

    Public Function GetElementByHandle(ByVal Hndl As IntPtr) As AutomationElement

        Try
            Return AutomationElement.FromHandle(Hndl)
        Catch ex As Exception
            Debug.Print(ex.Message)
            Return Nothing
        End Try

    End Function
    Public Function GetElementByName(ByRef window As AutomationElement, name As String) As AutomationElement
        If VarType(window) = vbNull Then
            Return Nothing
        End If
        Try
            Return window.FindFirst(TreeScope.Descendants, New PropertyCondition(AutomationElement.NameProperty, name))
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Function GetElementByClass(ByRef window As AutomationElement, classname As String) As AutomationElement
        If VarType(window) = vbNull Then
            Return Nothing
        End If
        Try
            Return window.FindFirst(TreeScope.Children, New PropertyCondition(AutomationElement.ClassNameProperty, classname))
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Function GetElementById(ByRef window As AutomationElement, id As String) As AutomationElement
        If IsNothing(window) Then
            Return Nothing
        End If
        Return window.FindFirst(TreeScope.Descendants, New PropertyCondition(AutomationElement.AutomationIdProperty, id))
    End Function

    Public Function InvokeElementByName(ByRef window As AutomationElement, name As String) As Boolean
        Return InvokeElement(window, New PropertyCondition(AutomationElement.NameProperty, name))
    End Function

    Public Function InvokeElementById(ByRef window As AutomationElement, id As String) As Boolean
        Return InvokeElement(window, New PropertyCondition(AutomationElement.AutomationIdProperty, id))
    End Function

    Public Function InvokeElement(ByRef window As AutomationElement, _
                                  condition As System.Windows.Automation.Condition) As Boolean

        If window Is Nothing Then Return False

        Dim element As AutomationElement = window.FindFirst(TreeScope.Descendants, condition)
        If Not IsNothing(element) Then
            Dim pattern As InvokePattern = element.GetCurrentPattern(InvokePattern.Pattern)
            pattern.Invoke()
            Return True
        End If
        Return False
    End Function
    Private Function FindMenu(ByRef window As AutomationElement, ByVal menu As String) As ExpandCollapsePattern
        Dim menuElement As AutomationElement = window.FindFirst(TreeScope.Descendants, _
                                                                New PropertyCondition(AutomationElement.NameProperty, menu))
        Dim expPattern As ExpandCollapsePattern = menuElement.GetCurrentPattern(ExpandCollapsePattern.Pattern)
        Return expPattern
    End Function

    Public Sub OpenMenu(ByRef window As AutomationElement, ByVal menu As String)
        Dim expPattern As ExpandCollapsePattern = FindMenu(window, menu)
        expPattern.Expand()
    End Sub

    Public Sub CloseMenu(ByRef window As AutomationElement, ByVal menu As String)
        Dim expPattern As ExpandCollapsePattern = FindMenu(window, menu)
        expPattern.Collapse()
    End Sub

    Public Sub ExecuteMenuByName(ByRef window As AutomationElement, ByVal menuName As String)
        Dim menuElement As AutomationElement = window.FindFirst(TreeScope.Descendants, New PropertyCondition(AutomationElement.NameProperty, menuName))
        If Not IsNothing(menuElement) Then
            Dim invokePattern As InvokePattern = menuElement.GetCurrentPattern(invokePattern.Pattern)
            If Not IsNothing(invokePattern) Then
                invokePattern.Invoke()
            End If
        End If
    End Sub


    Public Sub ExecuteMenuItem(ByRef window As AutomationElement, ByVal menuItem As String)
        Dim menuItemAndCondition As AndCondition = New AndCondition(New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem), _
                                                               New PropertyCondition(AutomationElement.NameProperty, menuItem))
        Dim menuItemElement As AutomationElement = window.FindFirst(TreeScope.Descendants, menuItemAndCondition)
        If Not IsNothing(menuItemElement) Then
            Dim invokePattern As InvokePattern = menuItemElement.GetCurrentPattern(invokePattern.Pattern)
            If Not IsNothing(invokePattern) Then
                invokePattern.Invoke()
            End If
        End If
    End Sub
    Public Function ClickMe(ByRef control As AutomationElement) As Boolean
        Try
            '' Need to check to see if it supports InvokePattern

            Dim selectionItem As InvokePattern = Nothing ''= control.GetCurrentPattern(InvokePattern.Pattern)
            If control.TryGetCurrentPattern(InvokePattern.Pattern, selectionItem) Then
                ''If Not IsNothing(selectionItem) Then
                selectionItem = control.GetCurrentPattern(InvokePattern.Pattern)
                selectionItem.Invoke()
                Return True
            End If
            '' If not lets see if it supports collapsing/expanding
            Dim expandCollapse As ExpandCollapsePattern = Nothing ''= control.GetCurrentPattern(ExpandCollapsePattern.Pattern)
            If control.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, expandCollapse) Then
                '' If Not IsNothing(expandCollapse) Then
                expandCollapse = control.GetCurrentPattern(ExpandCollapsePattern.Pattern)
                expandCollapse.Expand()
                Return True
            End If
        Catch ex As Exception
            Debug.Print("ClickMe event : " & control.Current.Name & ": " & ex.Message)
            Return False
        End Try
        Return False
    End Function
    ''' <summary>
    ''' Examples of using predefined conditions to find elements.
    ''' </summary>
    ''' <param name="elementMainWindow">The element for the target window.</param>
    Public Sub DumpControlNameId(ByVal elementMainWindow As AutomationElement)
        If elementMainWindow Is Nothing Then
            Throw New ArgumentException()
        End If

        ' Use TrueCondition to retrieve all elements.
        Dim elementCollectionAll As AutomationElementCollection = elementMainWindow.FindAll(TreeScope.Subtree, Condition.TrueCondition)
        Console.WriteLine(vbLf + "All control types:")
        Console.WriteLine("------------------------------")
        Dim autoElement As AutomationElement
        For Each autoElement In elementCollectionAll
            Console.WriteLine(autoElement.Current.Name)
            Console.WriteLine("  AutomationId = " & autoElement.Current.AutomationId.ToString)
        Next autoElement

        ' Use ContentViewCondition to retrieve all content elements.
        Dim elementCollectionContent As AutomationElementCollection = elementMainWindow.FindAll(TreeScope.Subtree, Automation.ContentViewCondition)
        Console.WriteLine(vbLf + "All content elements:")
        Console.WriteLine("------------------------------")
        For Each autoElement In elementCollectionContent
            Console.WriteLine(autoElement.Current.Name)
            Console.WriteLine("  AutomationId = " & autoElement.Current.AutomationId.ToString)
        Next autoElement

        ' Use ControlViewCondition to retrieve all control elements.
        Dim elementCollectionControl As AutomationElementCollection = elementMainWindow.FindAll(TreeScope.Subtree, Automation.ControlViewCondition)
        Console.WriteLine(vbLf & "All control elements:")
        Console.WriteLine("------------------------------")
        For Each autoElement In elementCollectionControl
            Console.WriteLine(autoElement.Current.Name)
            Console.WriteLine("  AutomationId = " & autoElement.Current.AutomationId.ToString)

        Next autoElement

    End Sub 'StaticConditionExamples
    Public Function GetElementContainsName(ByRef window As AutomationElement, strlookup As String) As AutomationElement
        Try
            Dim elementCollectionAll As AutomationElementCollection = window.FindAll(TreeScope.Subtree, Condition.TrueCondition)
            Dim autoElement As AutomationElement
            For Each autoElement In elementCollectionAll
                ''Debug.Print("Name = " & autoElement.Current.Name & " ID = " & autoElement.Current.AutomationId.ToString())
                '' If we have an AutomationId too then we will return the Element
                If InStr(autoElement.Current.Name, strlookup, CompareMethod.Binary) > 0 And Not String.IsNullOrEmpty(autoElement.Current.AutomationId.ToString) Then
                    ''we found it!
                    Return autoElement
                End If
            Next
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Sub CloseWindow(ByRef window As AutomationElement)
        Dim windowUI As WindowPattern = Nothing 'window.GetCurrentPattern(WindowPattern.Pattern)
        Try
            windowUI = DirectCast(window.GetCurrentPattern(WindowPattern.Pattern), WindowPattern)
            If True = windowUI.WaitForInputIdle(500) Then
                windowUI.Close()
            End If
        Catch ex As Exception
            Debug.Print("CloseWindow error " & ex.Message)
        End Try

    End Sub

    ''' --------------------------------------------------------------------
    '''  <summary>
    ''' Sets the values of the text controls using managed methods.
    ''' </summary>
    '''  <param name="s">The string to be inserted.</param>
    '''--------------------------------------------------------------------
    Public Sub SetValueWithUIAutomation(ByRef control As AutomationElement, ByVal s As String)

        InsertTextWithUIAutomation(control, s)


    End Sub


    '' --------------------------------------------------------------------
    ''  
    ''  Code taken from msdn with some custom tweaks
    '''<summary>
    '''  Inserts a string into each text control of interest.
    '''  </summary>
    '''  <param name="element">A text control.</param>
    '''  <param name="value">The string to be inserted.</param>
    ''' --------------------------------------------------------------------
    Private Sub InsertTextWithUIAutomation( _
    ByVal element As AutomationElement, ByVal value As String)
        Try
            ' Validate arguments / initial setup
            If value Is Nothing Then
                Throw New ArgumentNullException( _
                "String parameter must not be null.")

            End If

            If element Is Nothing Then
                Throw New ArgumentNullException( _
                "AutomationElement parameter must not be null")
            End If

            ' A series of basic checks prior to attempting an insertion.
            '
            ' Check #1: Is control enabled?
            ' An alternative to testing for static or read-only controls 
            ' is to filter using 
            ' PropertyCondition(AutomationElement.IsEnabledProperty, true) 
            ' and exclude all read-only text controls from the collection.
            If Not element.Current.IsEnabled Then
                Throw New InvalidOperationException( _
                "The control with an AutomationID of " + _
                element.Current.AutomationId.ToString() + _
                " is not enabled.")
            End If

            ' Check #2: Are there styles that prohibit us 
            '           from sending text to this control?
            If Not element.Current.IsKeyboardFocusable Then
                Throw New InvalidOperationException( _
                "The control with an AutomationID of " + _
                element.Current.AutomationId.ToString() + _
                "is read-only.")
            End If


            ' Once you have an instance of an AutomationElement,  
            ' check if it supports the ValuePattern pattern.
            Dim targetValuePattern As Object = Nothing

            ' Control does not support the ValuePattern pattern 
            ' so use keyboard input to insert content.
            '
            ' NOTE: Elements that support TextPattern 
            '       do not support ValuePattern and TextPattern
            '       does not support setting the text of 
            '       multi-line edit or document controls.
            '       For this reason, text input must be simulated
            '       using one of the following methods.
            '       
            If Not element.TryGetCurrentPattern(ValuePattern.Pattern, targetValuePattern) Then
                Debug.Print("The control with an AutomationID of ")
                Debug.Print(element.Current.AutomationId.ToString())
                Debug.Print(" does not support ValuePattern.")
                Debug.Print(" Using keyboard input.")

                ' Set focus for input functionality and begin.
                element.SetFocus()

                ' Pause before sending keyboard input.
                Thread.Sleep(100)

                ' Delete existing content in the control and insert new content.
                SendKeys.SendWait("^{HOME}") ' Move to start of control
                SendKeys.SendWait("^+{END}") ' Select everything
                SendKeys.SendWait("{DEL}") ' Delete selection
                SendKeys.SendWait(value)
            Else
                ' Control supports the ValuePattern pattern so we can 
                ' use the SetValue method to insert content.
                Debug.Print("The control with an AutomationID of ")
                Debug.Print(element.Current.AutomationId.ToString())
                Debug.Print(" supports ValuePattern.")
                Debug.Print(" Using ValuePattern.SetValue().")

                ' Set focus for input functionality and begin.
                element.SetFocus()
                Dim valueControlPattern As ValuePattern = _
                DirectCast(targetValuePattern, ValuePattern)
                valueControlPattern.SetValue(value)
            End If
        Catch exc As ArgumentNullException
            Debug.Print(exc.Message)
        Catch exc As InvalidOperationException
            Debug.Print(exc.Message)
        Finally

        End Try

    End Sub

    Public Function TraverseMenu(ByRef window As AutomationElement, ByVal menulist As List(Of String)) As Boolean
        Dim i As Int64
        If menulist.Count = 0 Then Return False
        Try
            If menulist.Count = 1 Then
                '' Only one menu Item
                OpenMenu(window, menulist(0))
                Return True
            Else
                '' First OpenMenu and loop through the List of menu items
                OpenMenu(window, menulist(0))
                For i = 1 To menulist.Count - 1
                    ExecuteMenuItem(window, menulist(i))
                Next

            End If
        Catch e As Exception
            Debug.Print(e.Message)
            Return False
        End Try
        Return True
    End Function

    Public Function GetElementText(ByRef window As AutomationElement)
        Try
            If Not IsNothing(window) Then
                Return window.Current.Name
            End If
        Catch ex As Exception
            Debug.Print(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Sub Focus(ByRef window As AutomationElement)
        Try
            If Not IsNothing(window) Then
                window.SetFocus()
            End If
        Catch ex As Exception
            Debug.Print("Focus: " & window.Current.Name)
            Debug.Print(vbTab & ex.Message)

        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
