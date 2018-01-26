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
                    txtModsDir.Text = modPath
                    My.Settings.modsFolderPath = modPath
                    My.Settings.Save()
                End If
            End If
        Else
                Dim ask As MsgBoxResult = MsgBox("A mods folder doesn't exist yet in this location. Would you like the mod manager to create it for you?", MsgBoxStyle.YesNo, "Create Mods Folder")
            If ask = MsgBoxResult.Yes Then
                MkDir(modPath)
                txtModsDir.Text = modPath
                My.Settings.modsFolderPath = modPath
                My.Settings.Save()
            Else
                MsgBox("OK. In that case, please specify the mods folder manually.", MsgBoxStyle.OkOnly, "Create Mods Folder")
            End If
        End If
    End Sub
#End Region

#Region "Mods Folder Path"
    ' Read the property of the User Setting detailing the Mods Folder Path, and put it as the value for the text box
    Private Sub setModsDir(sender As Object, e As EventArgs) Handles txtModsDir.Initialized
        txtModsDir.Text = My.Settings.modsFolderPath
    End Sub

    ' Clicked the Change Directory button associated with the Mods Directory
    Private Sub ModsFolderDirChangeBtn_Click(sender As Object, e As RoutedEventArgs) Handles modsFolderDirChangeBtn.Click
        ModsFolderChange()
    End Sub

    ' Open a directory browser that allows you to select a folder, then save the selection to a user setting to store the selected folder for later use
    Private Sub ModsFolderChange()
        ' When this is clicked, open a directory dialog
        Dim modsFolderPicker As New CommonOpenFileDialog With {
            .IsFolderPicker = True
        }
        modsFolderPicker.ShowDialog()
        txtModsDir.Text = modsFolderPicker.FileName
        My.Settings.modsFolderPath = modsFolderPicker.FileName
        My.Settings.Save()
    End Sub

#End Region



#Region "Menu Bar"
#Region "File Menu"
    ' Open the directory browser to change the Mods folder when File > Change Mods Directory is clicked
    Private Sub MenuFile_ChangeModsDir_Click(sender As Object, e As RoutedEventArgs)
        ModsFolderChange()
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
