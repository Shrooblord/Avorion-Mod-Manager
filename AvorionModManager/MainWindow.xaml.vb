Imports System.IO
Imports Microsoft.WindowsAPICodePack.Dialogs

Class MainWindow
#Region "Avorion Installation Directory"
    ' Read the property of the User Setting detailing the Avorion Installation Path, and put it as the value for the text box
    Private Sub setAvoDir(sender As Object, e As EventArgs) Handles txtAvoDir.Initialized
        txtAvoDir.Text = My.Settings.avorionInstallPath
    End Sub

    ' Clicked the Change Directory button associated with the Avorion Installation Directory
    Private Sub AvorionDirChangeBtn_Click(sender As Object, e As RoutedEventArgs) Handles avorionDirChangeBtn.Click
        AvoInstallDirChange()
    End Sub

    ' Open a directory browser that allows you to select a folder, then save the selection to a user setting to store the selected folder for later use
    Private Sub AvoInstallDirChange()
        ' When this is clicked, open a directory dialog
        Dim avoDirPicker As New CommonOpenFileDialog With {
            .IsFolderPicker = True
        }
        avoDirPicker.ShowDialog()
        txtAvoDir.Text = avoDirPicker.FileName
        My.Settings.avorionInstallPath = avoDirPicker.FileName
        My.Settings.Save()

        CheckModsFolderExists(avoDirPicker.FileName)
    End Sub

    ' Check to see if a mods folder exists in the location of the Avorion Installation Directory, and if so, if the mod manager should point to it.
    ' If not, ask whether it should be created automatically.
    Private Sub CheckModsFolderExists(path As String)
        Dim modPath As String
        modPath = path & "\mods"
        If Directory.Exists(modPath) Then
            ' Only ask this if the Mods Folder User Setting doesn't already point to the folder we found
            If Not My.Settings.modsFolderPath = modPath Then
                Dim ask As MsgBoxResult = MsgBox("A mods folder was found in this location. Would you like the mod manager to use this folder?", MsgBoxStyle.YesNo, "Select Mods Folder")
                If ask = MsgBoxResult.Yes Then
                    ModsFolderChange(modPath)
                End If
            End If
        Else
                Dim ask As MsgBoxResult = MsgBox("A mods folder doesn't exist yet in this location. Would you like the mod manager to create it for you?", MsgBoxStyle.YesNo, "Create Mods Folder")
            If ask = MsgBoxResult.Yes Then
                MkDir(modPath)
                ModsFolderChange(modPath)
            Else
                MsgBox("OK. In that case, please specify the mods folder manually.", MsgBoxStyle.OkOnly, "Create Mods Folder")
            End If
        End If
    End Sub
#End Region

#Region "Mods Folder Path"
    ' Read the property of the User Setting detailing the Mods Folder Path, and put it as the value for the text box
    Private Sub SetModsDir(sender As Object, e As EventArgs) Handles txtModsDir.Initialized
        txtModsDir.Text = My.Settings.modsFolderPath
    End Sub

    ' Clicked the Change Directory button associated with the Mods Directory
    Private Sub ModsFolderDirChangeBtn_Click(sender As Object, e As RoutedEventArgs) Handles modsFolderDirChangeBtn.Click
        ModsFolderPicker()
    End Sub

    ' Open a directory browser that allows you to select a folder, then save the selection to a user setting to store the selected folder for later use
    Private Sub ModsFolderPicker()
        ' When this is clicked, open a directory dialog
        Dim modsFolderPicker As New CommonOpenFileDialog With {
            .IsFolderPicker = True
        }
        modsFolderPicker.ShowDialog()
        ModsFolderChange(modsFolderPicker.FileName)
    End Sub

    Private Sub ModsFolderChange(modDir As String)
        txtModsDir.Text = modDir
        My.Settings.modsFolderPath = modDir
        My.Settings.Save()
        ' Populate the Uninstalled Mods Listbox
        AddAllFoldersToUninstalledListBox(modDir)
    End Sub

#End Region

#Region "Listbox Uninstalled Mods"
    ' Find all folders inside the Mods Directory and put them in the Uninstalled Mods Listbox
    Private Sub AddAllFoldersToUninstalledListBox(modDir As String)
        For Each Dir As String In Directory.GetDirectories(modDir)
            ' Put it as a list entry in the Uninstalled Mods Listbox
            Dim modName As String
            modName = Dir
            ' Replaces the string modDir with an empty string, effectively cutting out the directory leading up to the speficic mod's folder, and leaving
            ' only the folder's name i.e. the mod's name to be added to the listbox
            modName = modName.Replace(modDir & "\", "")
            listboxUninstalledMods.Items.Add(modName)
        Next
    End Sub

    ' Attempt to populate the list; see if the mods folder has already been set
    Private Sub InitListboxUninstalledMods(sender As Object, e As EventArgs) Handles listboxUninstalledMods.Initialized
        If My.Settings.modsFolderPath <> String.Empty Then
            AddAllFoldersToUninstalledListBox(My.Settings.modsFolderPath)
        End If
    End Sub
#End Region

#Region "Menu Bar"
#Region "File Menu"
    ' Open the directory browser to change the Mods folder when File > Change Mods Directory is clicked
    Private Sub MenuFile_ChangeModsDir_Click(sender As Object, e As RoutedEventArgs)
        ModsFolderPicker()
    End Sub

    ' Open the directory browser to change the Avorion Installation directory when File > Change Avorion Installation Directory is clicked
    Private Sub MenuFile_ChangeAvoDir_Click(sender As Object, e As RoutedEventArgs)
        AvoInstallDirChange()
    End Sub

    ' Exit the program when File > Exit is clicked
    Private Sub MenuFile_Exit_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub
#End Region
#End Region
End Class
