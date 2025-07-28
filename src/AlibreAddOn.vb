Imports AlibreAddOn
Imports AlibreX
Imports IronPython.Hosting
Imports Microsoft.Scripting.Hosting
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Windows
Imports IStream = System.Runtime.InteropServices.ComTypes.IStream
Imports MessageBox = System.Windows.MessageBox

Namespace AlibreAddOnAssembly
    Public Module AlibreAddOn
        Private Property AlibreRoot As IADRoot
        Private Property TheAddOnInterface As AddOnRibbon
        Private Property PythonRunner As ScriptRunner

        Public Sub AddOnLoad(ByVal hwnd As IntPtr, ByVal pAutomationHook As IAutomationHook, ByVal unused As IntPtr)
            AlibreRoot = CType(pAutomationHook.Root, IADRoot)
            PythonRunner = New ScriptRunner(AlibreRoot)
            TheAddOnInterface = New AddOnRibbon(AlibreRoot)
        End Sub

        Public Sub AddOnUnload(ByVal hwnd As IntPtr, ByVal forceUnload As Boolean, ByRef cancel As Boolean, ByVal reserved1 As Integer, ByVal reserved2 As Integer)
            TheAddOnInterface = Nothing
            PythonRunner = Nothing
            AlibreRoot = Nothing
        End Sub

        Public Sub AddOnInvoke(ByVal pAutomationHook As IntPtr, ByVal sessionName As String, ByVal isLicensed As Boolean, ByVal reserved1 As Integer, ByVal reserved2 As Integer)
        End Sub

        Public Function GetAddOnInterface() As IAlibreAddOn
            Return TheAddOnInterface
        End Function

        Public Function GetScriptRunner() As ScriptRunner
            Return PythonRunner
        End Function
    End Module

    Public Class AddOnRibbon
        Implements IAlibreAddOn

        Private ReadOnly _menuManager As MenuManager
        Private ReadOnly _alibreRoot As IADRoot

        Public Sub New(ByVal alibreRoot As IADRoot)
            _alibreRoot = alibreRoot
            _menuManager = New MenuManager()
        End Sub

        Public ReadOnly Property RootMenuItem As Integer Implements IAlibreAddOn.RootMenuItem
            Get
                Return _menuManager.GetRootMenuItem().Id
            End Get
        End Property

        Public Function HasSubMenus(ByVal menuID As Integer) As Boolean Implements IAlibreAddOn.HasSubMenus
            Dim menuItem = _menuManager.GetMenuItemById(menuID)
            Return If(menuItem?.SubItems.Count > 0, True, False)
        End Function

        Public Function SubMenuItems(ByVal menuID As Integer) As Array Implements IAlibreAddOn.SubMenuItems
            Dim menuItem = _menuManager.GetMenuItemById(menuID)
            Return menuItem?.SubItems.Select(Function(subItem) subItem.Id).ToArray()
        End Function

        Public Function MenuItemText(ByVal menuID As Integer) As String Implements IAlibreAddOn.MenuItemText
            Return _menuManager.GetMenuItemById(menuID)?.Text
        End Function

        Public Function MenuItemToolTip(ByVal menuID As Integer) As String Implements IAlibreAddOn.MenuItemToolTip
            Return _menuManager.GetMenuItemById(menuID)?.ToolTip
        End Function

        ' Icon functionality disabled: always returns Nothing
        Public Function MenuIcon(ByVal menuID As Integer) As String Implements IAlibreAddOn.MenuIcon
            Return Nothing
        End Function

        Public Function InvokeCommand(ByVal menuID As Integer, ByVal sessionIdentifier As String) As IAlibreAddOnCommand Implements IAlibreAddOn.InvokeCommand
            Dim session = _alibreRoot.Sessions.Item(sessionIdentifier)
            Dim menuItem = _menuManager.GetMenuItemById(menuID)
            Return menuItem?.Command?.Invoke(session)
        End Function

        Public Function MenuItemState(ByVal menuID As Integer, ByVal sessionIdentifier As String) As ADDONMenuStates Implements IAlibreAddOn.MenuItemState
            Return ADDONMenuStates.ADDON_MENU_ENABLED
        End Function

        Public Function PopupMenu(ByVal menuID As Integer) As Boolean Implements IAlibreAddOn.PopupMenu
            Return False
        End Function

        Public Function HasPersistentDataToSave(ByVal sessionIdentifier As String) As Boolean Implements IAlibreAddOn.HasPersistentDataToSave
            Return False
        End Function

        Public Sub SaveData(ByVal pCustomData As Global.AlibreAddOn.IStream, ByVal sessionIdentifier As String) Implements IAlibreAddOn.SaveData
        End Sub

        Public Sub LoadData(ByVal pCustomData As Global.AlibreAddOn.IStream, ByVal sessionIdentifier As String) Implements IAlibreAddOn.LoadData
        End Sub

        Public Function UseDedicatedRibbonTab() As Boolean Implements IAlibreAddOn.UseDedicatedRibbonTab
            Return False
        End Function

        Public Sub setIsAddOnLicensed(ByVal isLicensed As Boolean) Implements IAlibreAddOn.setIsAddOnLicensed
        End Sub
    End Class

    Public Class MenuItem
        Public Property Id As Integer
        Public Property Text As String
        Public Property ToolTip As String
        Public Property Icon As String
        Public Property Command As Func(Of IADSession, IAlibreAddOnCommand)
        Public Property SubItems As List(Of MenuItem) = New List(Of MenuItem)()

        Public Sub New(ByVal id As Integer, ByVal text As String, Optional ByVal toolTip As String = "", Optional ByVal icon As String = Nothing)
            Me.Id = id
            Me.Text = text
            Me.ToolTip = toolTip
            Me.Icon = Nothing
        End Sub

        Public Sub AddSubItem(ByVal subItem As MenuItem)
            SubItems.Add(subItem)
        End Sub

        Public Function AboutCmd(ByVal session As IADSession) As IAlibreAddOnCommand
            MessageBox.Show("Skinning add-on demo" & vbCrLf & vbCrLf)
            Return Nothing
        End Function
    End Class

    Public Class MenuManager
        Private ReadOnly _rootMenuItem As MenuItem
        Private ReadOnly _menuItems As Dictionary(Of Integer, MenuItem) = New Dictionary(Of Integer, MenuItem)()

        Public Sub New()
            _rootMenuItem = New MenuItem(401, "alibre-skinning-addon", "alibre-skinning-addon")
            BuildMenus()
            RegisterMenuItem(_rootMenuItem)
        End Sub

        Private Sub BuildMenus()
            Dim aboutItem = New MenuItem(9090, "About", "https://github.com/stephensmitchell/alibre-skinning-addon")
            aboutItem.Command = AddressOf aboutItem.AboutCmd
            _rootMenuItem.AddSubItem(aboutItem)

            Try
                Dim addOnDirectory As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                Dim examplesPath As String = Path.Combine(addOnDirectory, "Scripts\src\hss")
                If Directory.Exists(examplesPath) Then
                    Dim currentMenuId As Integer = 10000
                    Dim scriptFiles = Directory.GetFiles(examplesPath, "*.py")
                    For Each scriptFile In scriptFiles
                        Dim fileName As String = Path.GetFileName(scriptFile)
                        If fileName.Equals("alibre_setup.py", StringComparison.OrdinalIgnoreCase) Then Continue For
                        Dim baseName As String = Path.GetFileNameWithoutExtension(fileName)
                        Dim menuText As String = baseName.Replace("-", " ").Replace("_", " ")
                        Dim scriptMenuItem = New MenuItem(currentMenuId, menuText, $"Run {fileName}")
                        currentMenuId += 1
                        scriptMenuItem.Command = Function(session)
                                                     AlibreAddOn.GetScriptRunner()?.ExecuteScript(session, fileName)
                                                     Return Nothing
                                                 End Function
                        _rootMenuItem.AddSubItem(scriptMenuItem)
                    Next
                End If
            Catch ex As Exception
                MessageBox.Show($"Failed to load scripts dynamically: {ex.Message}", "Add-on Error")
            End Try
        End Sub

        Private Sub RegisterMenuItem(ByVal menuItem As MenuItem)
            _menuItems(menuItem.Id) = menuItem
            For Each subItem In menuItem.SubItems
                RegisterMenuItem(subItem)
            Next
        End Sub

        Public Function GetMenuItemById(ByVal id As Integer) As MenuItem
            Dim menuItem As MenuItem = Nothing
            Return If(_menuItems.TryGetValue(id, menuItem), menuItem, Nothing)
        End Function

        Public Function GetRootMenuItem() As MenuItem
            Return _rootMenuItem
        End Function
    End Class

    Public Class ScriptRunner
        Private ReadOnly _engine As ScriptEngine
        Private ReadOnly _alibreRoot As IADRoot

        Public Sub New(ByVal alibreRoot As IADRoot)
            _alibreRoot = alibreRoot
            _engine = Python.CreateEngine()
            Dim alibreInstallPath As String = "C:\Program Files\Alibre Design 28.1.1.28227"
            Dim searchPaths = _engine.GetSearchPaths()
            searchPaths.Add(Path.Combine(alibreInstallPath, "Program"))
            searchPaths.Add(Path.Combine(alibreInstallPath, "Program", "Addons", "AlibreScript", "PythonLib"))
            searchPaths.Add(Path.Combine(alibreInstallPath, "Program", "Addons", "AlibreScript"))
            _engine.SetSearchPaths(searchPaths)
        End Sub

        Public Sub ExecuteScript(ByVal session As IADSession, ByVal mainScriptFileName As String)
            Try
                Dim addOnDirectory As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                Dim ScriptsPath As String = Path.Combine(addOnDirectory, "Scripts\src\hss")
                Dim setupScriptPath As String = Path.Combine(ScriptsPath, "alibre_setup.py")
                Dim mainScriptPath As String = Path.Combine(ScriptsPath, mainScriptFileName)

                If Not File.Exists(setupScriptPath) OrElse Not File.Exists(mainScriptPath) Then
                    MessageBox.Show($"Error: Script not found.{vbLf}Setup: {setupScriptPath}{vbLf}Main: {mainScriptPath}", "Script Error")
                    Return
                End If

                Dim scope As ScriptScope = _engine.CreateScope()
                scope.SetVariable("ScriptFileName", mainScriptFileName)
                scope.SetVariable("ScriptFolder", ScriptsPath)
                scope.SetVariable("SessionIdentifier", session.Identifier)
                scope.SetVariable("Arguments", New List(Of String)())
                scope.SetVariable("AlibreRoot", _alibreRoot)
                scope.SetVariable("CurrentSession", session)

                _engine.ExecuteFile(setupScriptPath, scope)
                _engine.ExecuteFile(mainScriptPath, scope)
            Catch ex As Exception
                MessageBox.Show($"An error occurred while running the script:{vbLf}{ex}", "Python Execution Error")
            End Try
        End Sub
    End Class
End Namespace
