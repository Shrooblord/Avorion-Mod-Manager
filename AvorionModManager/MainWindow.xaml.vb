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
        ' When this is clicked, open a directory dialog
        Dim avoDirPicker As New CommonOpenFileDialog With {
            .IsFolderPicker = True
        }
        avoDirPicker.ShowDialog()
        txtAvoDir.Text = avoDirPicker.FileName
        My.Settings.avorionInstallPath = avoDirPicker.FileName
        My.Settings.Save()
    End Sub
#End Region

#Region "Mods Folder Path"
    ' Read the property of the User Setting detailing the Mods Folder Path, and put it as the value for the text box
    Private Sub setModsDir(sender As Object, e As EventArgs) Handles txtModsDir.Initialized
        txtModsDir.Text = My.Settings.modsFolderPath
    End Sub

    ' Clicked the Change Directory button associated with the Mods Directory
    Private Sub ModsFolderDirChangeBtn_Click(sender As Object, e As RoutedEventArgs) Handles modsFolderDirChangeBtn.Click
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
    ' Exit the program when File > Exit is clicked
    Private Sub MenuFile_Exit_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub
#End Region
#End Region
End Class
