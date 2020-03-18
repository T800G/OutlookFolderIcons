Imports Microsoft.Office.Interop
Imports Microsoft.Interop.Stdole
Imports System.Windows.Forms

Public Class FolderIcons
	Private m_settings As xmlSettings
	Private m_app As Outlook.Application
	Private m_iconSize as Int32
	
	Public ReadOnly Property ImageLibraryFolder() As String
		Get
			return m_settings.ImageLibraryFolder
		End Get
	End Property
	
	Public ReadOnly Property iconSize() As Int32
		Get
			return m_iconSize
		End Get
	End Property
		
	Public Sub Initialize(App As Object)
		If m_app Is Nothing Then m_app = CType(App, Outlook.Application)
		m_iconSize = System.Windows.Forms.SystemInformation.SmallIconSize.Width 'SM_CXSMICON

		'Outlook uses 16x16 icons for folder treeview while dpi is under 125% but their width is distorted
		If m_iconSize < 21 Then m_iconSize = 16

		If m_settings Is Nothing Then m_settings = New xmlSettings
		Dim oStore As Outlook.Store
		For Each oStore In m_app.Session.Stores
			EnumerateFolders(oStore.GetRootFolder())
		Next
	End Sub
	
	Private Sub EnumerateFolders(ByRef oFolder As Outlook.MAPIFolder)
		SetIconImpl(oFolder)
	    If oFolder.Folders.Count > 0 Then
	        Dim fldr As Outlook.MAPIFolder
	        For Each fldr In oFolder.Folders
	            EnumerateFolders(fldr) 'recursion!!!
	        Next
	    End If
	End Sub

	Public Function IsCustomized(Optional oFolder As Outlook.MAPIFolder = Nothing) As Boolean
		If oFolder Is Nothing Then
			If m_app.ActiveExplorer Is Nothing Then Return True
			oFolder = m_app.ActiveExplorer.CurrentFolder
		End If
	    If oFolder.GetCustomIcon() Is Nothing Then Return False
	    IsCustomized = True
	End Function
	
	Public Function IsDefaultFolder(Optional oFolder As Outlook.MAPIFolder = Nothing) As Boolean
		If oFolder Is Nothing Then
			If m_app.ActiveExplorer Is Nothing Then Return True
			oFolder = m_app.ActiveExplorer.CurrentFolder
		End If
		If IsRootFolder(oFolder) Or IsDefaultFolderImpl(oFolder) Or IsSpecialFolder(ofolder) Then Return True
		Return False
	End Function
	
	Private Function IsRootFolder(oFolder As Outlook.MAPIFolder) As Boolean
		Dim f As Outlook.MAPIFolder = Nothing
		Try: f = oFolder.Store.GetRootFolder()
		Catch ex As System.Exception
		End Try
		If Not f Is Nothing Then
			If (oFolder.entryID = f.entryID) Then Return True
			f = Nothing
		End If
		Return False
	End Function
	
	Private Function IsSpecialFolder(oFolder As Outlook.MAPIFolder) As Boolean
		For Each i As Outlook.OlSpecialFolders In System.Enum.GetValues(GetType(Outlook.OlSpecialFolders))
			Dim f As Outlook.MAPIFolder = Nothing
			Try: f = oFolder.Store.GetSpecialFolder(i) 'not all special folders might exist in a store!
			Catch ex As System.Exception
			End Try
			If Not f Is Nothing Then
				If (oFolder.entryID = f.entryID) Then Return True
				f = Nothing
			End If
		next
		Return False
	End Function
	
	Private Function IsDefaultFolderImpl(oFolder As Outlook.MAPIFolder) As Boolean
		For Each i As Outlook.OlDefaultFolders In System.Enum.GetValues(GetType(Outlook.OlDefaultFolders))
			Dim f As Outlook.MAPIFolder = Nothing
			Try: f = oFolder.Store.GetDefaultFolder(i) 'not all default folders might exist in a store!
			Catch ex As System.Exception
			End Try
			If Not f Is Nothing Then
				If (oFolder.entryID = f.entryID) Then Return True
				f = Nothing
			End If
			Next
			Return False
	End Function

	Public Sub RemoveIcon(Optional oFolder As Outlook.MAPIFolder = Nothing)
		If oFolder Is Nothing Then
			If m_app.ActiveExplorer Is Nothing Then Exit Sub
			oFolder = m_app.ActiveExplorer.CurrentFolder
		End If
	    If oFolder.GetCustomIcon() Is Nothing Then Exit Sub 'this also handles default/special folders
	    If MessageBox.Show("Restore default icon for " & oFolder.name & "?", "Folder Icons", _
	    					MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2 _
	    					) = DialogResult.Cancel Then Exit Sub
			Try: oFolder.SetCustomIcon(nothing) 'undocumented way to get default icon back!
			Catch ex As System.Exception
				'trace.TraceError(ex.ToString())
			End Try
			m_settings.DeleteSetting(oFolder.StoreID, oFolder.EntryID)
	End Sub
	
	Public Sub SetIcon(Optional oFolder As Outlook.MAPIFolder = Nothing, Optional iconPath As String = vbnullstring)
		If oFolder Is Nothing Then
			If m_app.ActiveExplorer Is Nothing Then Exit Sub
			oFolder = m_app.ActiveExplorer.CurrentFolder
		End If
	    If IsDefaultFolder(oFolder) Then
	    	MessageBox.Show("Custom icon cannot be set for a default or special folder", "Folder Icons", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
	        Exit Sub
	    End If
	    If iconPath = vbnullstring then 
		    Dim ofd As New System.Windows.Forms.OpenFileDialog
		    ofd.InitialDirectory = m_settings.ImageLibraryFolder
			ofd.CheckFileExists = True
			ofd.CheckPathExists = True
			ofd.Multiselect = False
			ofd.DereferenceLinks = True
			ofd.Title = "Select icon"
			ofd.Filter = "Icons (*.ico)|*.ico"
			If ofd.ShowDialog() <> System.Windows.Forms.DialogResult.OK then Exit Sub
			iconPath = ofd.FileName
		End if
		SetIconImpl(oFolder, iconPath)
	End Sub

	Private Sub SetIconImpl(oFolder As Outlook.MAPIFolder, Optional iconPath As String = vbnullstring)
		'called by startup enumerator (iconPath empty) and by SetIcon
		Dim path As String = iconPath
	    If path = "" Then path = m_settings.GetIconPath(oFolder.StoreID, oFolder.EntryID)
	   	If path = "" Then exit sub
		If Not system.IO.File.Exists(path) Then
			MessageBox.Show("Icon not found" & vbCrlf & path, "Folder Icons", MessageBoxButtons.OK, MessageBoxIcon.Warning)
			m_settings.DeleteSetting(oFolder.storeID, oFolder.entryID)
	        exit sub
	    End If
		Try
			oFolder.SetCustomIcon(StdPictureConverter.ImageToIPicture(New System.Drawing.Icon(path, Me.iconSize, Me.iconSize).ToBitmap()))
			m_settings.SetIconPath(oFolder.StoreID, oFolder.EntryID, path)
		Catch ex As System.Exception
		     'trace.TraceError(ex.ToString())
		     exit sub
		End Try
	End Sub
	
End Class