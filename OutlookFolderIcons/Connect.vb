Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Extensibility14
Imports Microsoft.Office.Interop.Outlook
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Resources
imports System.Diagnostics

<ComVisible(True), Guid("AC92B228-C86E-4500-B1DE-D6E78D4CD094"), ProgId("FolderIconsAddin.Connect")> _
Public Class Connect
	Implements IDTExtensibility2
	Implements IRibbonExtensibility

	Private m_olApp As Microsoft.Office.Interop.Outlook.Application
	Private WithEvents m_olExplorers As Microsoft.Office.Interop.Outlook.Explorers
	Private WithEvents m_olActiveExplorer As Microsoft.Office.Interop.Outlook.Explorer
	Private m_ribbon As IRibbonUI
	Private m_ExplorerRibbonXML As String	
	Private m_foldericons As FolderIcons
	Private m_imagefiles() As String
	
'TODO localize menu items, message strings, help files
'TODO move messageboxes to separate class (separate string localization code)
'TODO detect installed/UI language << Application.LanguageSettings.LanguageID(msoLanguageIDUI)
'TODO load external localized files  ((strings<langID>.xml, ribbon<langID>.xml, index<langID>.html))
	
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Private Sub LoadRibbonXML()
		m_ExplorerRibbonXML = GetEmbeddedResource(System.Reflection.Assembly.GetExecutingAssembly, "OutlookFolderIcons.ExplorerRibbon.xml")
	End Sub
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Public Function GetEmbeddedResource(assembly As System.Reflection.Assembly, name As String) As String
		Dim buf As String = ""
        Using s As System.IO.Stream = assembly.GetManifestResourceStream(name)
            Using sr As New System.IO.StreamReader(s)
                buf = sr.ReadToEnd
                sr.Close()
            End Using
            s.Close()
        End Using
        Return buf
    End Function
    
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Private Sub Initialize(App As Object)
		'debug.WriteLine("Connect::Initialize")
		If m_foldericons Is Nothing Then m_foldericons = New FolderIcons
		m_foldericons.Initialize(m_olApp)
		Erase m_imagefiles
		'enumerate icons only at start for better performance 
		m_imagefiles = System.IO.Directory.GetFiles(m_foldericons.ImageLibraryFolder, "*.ico", System.IO.SearchOption.AllDirectories)
		Array.Sort(m_imagefiles)
	End Sub
	
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'IMPORTANT: SetEventHandler method doesn't work because GC destroys objects which were hooked,
	'we have to use m_olExplorers to lock reference to Explorers object and event handler functions
	Private Sub NewExplorerHandler(exp As Explorer) Handles m_olExplorers.NewExplorer
		'debug.WriteLine("m_olExplorers.NewExplorer")	
		if m_olActiveExplorer is nothing then m_olActiveExplorer = m_olApp.ActiveExplorer
		Initialize(m_olApp)
	End Sub
	
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Private Sub FolderSwitchHandler() Handles m_olActiveExplorer.FolderSwitch
		'debug.WriteLine("m_olActiveExplorer.FolderSwitch")
		m_ribbon.Invalidate
	End Sub

	'IDTExtensibility2 ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Private Sub OnConnection(App As Object, ConnectMode As ext_ConnectMode, AddInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection
		Try
			LoadRibbonXML()
			m_olApp = CType(App, Microsoft.Office.Interop.Outlook.Application)
			m_olExplorers = m_olApp.Explorers
			If m_olExplorers.Count > 0 Then
				m_olActiveExplorer = m_olApp.ActiveExplorer
			end if
			Initialize(m_olApp)
		Catch ex As System.Exception
			MessageBox.Show("Error loading Folder Icons Add-in" & vbcrlf & ex.ToString())
		End Try
	End Sub

	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Private Sub OnDisconnection(RemoveMode As ext_DisconnectMode, ByRef custom As System.Array) Implements IDTExtensibility2.OnDisconnection
		m_ribbon = Nothing
		m_olActiveExplorer = Nothing
		m_olExplorers = Nothing
		m_olApp = Nothing
	End Sub
	
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Private Sub OnStartupComplete(ByRef custom As System.Array) Implements IDTExtensibility2.OnStartupComplete
	End Sub
	Private Sub OnAddInsUpdate(ByRef custom As System.Array) Implements IDTExtensibility2.OnAddInsUpdate
	End Sub
	Private Sub OnBeginShutdown(ByRef custom As System.Array) Implements IDTExtensibility2.OnBeginShutdown
	End Sub
	

	'IRibbonExtensibility ++++++++++++++++++++++++++++++++++++++++++++++
	Private Function GetCustomUI(ribbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
		Select ribbonID
			Case "Microsoft.Outlook.Explorer": return m_ExplorerRibbonXML
		End Select
		GetCustomUI = "<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui""><ribbon></ribbon></customUI>"
	End Function
	
	'ribbon callbacks ++++++++++++++++++++++++++++++++++++++++++++++++++
	Public Sub OnFolderIconsRibbonLoad(ByVal ribbon As IRibbonUI)
		m_ribbon = ribbon
	End Sub

	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Public Sub ActionButtonClick(control As IRibbonControl)
		Dim ctxFolder As Folder = Nothing
		Try: ctxFolder = CType(control.Context, Folder)	'if called from folder treeview context menu
		Catch ex As System.Exception
		End Try
		Select Case control.Id
			Case "folderIcons.Explorer.Button.SetIcon", _
				"folderIcons.ContextMenu.Button.SetIcon": m_foldericons.SetIcon(ctxFolder)

			Case "folderIcons.Explorer.Button.RemoveIcon", _
				"folderIcons.ContextMenu.Button.RemoveIcon": m_foldericons.RemoveIcon(ctxFolder)
				
			Case "folderIcons.Explorer.Button.OpenLibraryFolder": OpenLibraryFolder()
		    Case "folderIcons.Explorer.Button.Help": OpenHelpFile()
		End Select
		m_ribbon.Invalidate
	End Sub
	
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Public Function GetControlVisible(control As IRibbonControl) As Boolean
		'debug.WriteLine("GetControlVisible:  control.id=" & control.Id)		
		Dim ctxFolder As Folder = Nothing
		Try: ctxFolder = CType(control.Context, Folder)
		Catch ex As System.Exception
		End Try
		Select Case control.Id
			Case "folderIcons.ContextMenu.DynamicGallery", _
				"folderIcons.Explorer.DynamicGallery", _
				"folderIcons.Explorer.Button.SetIcon", _
				"folderIcons.ContextMenu.Button.SetIcon": Return Not m_foldericons.IsDefaultFolder(ctxFolder)

			Case "folderIcons.Explorer.Button.RemoveIcon", _
				"folderIcons.ContextMenu.Button.RemoveIcon": Return m_foldericons.IsCustomized(ctxFolder)

			Case "folderIcons.Explorer.CmdUnavailable", _
				"folderIcons.ContextMenu.CmdUnavailable": Return m_foldericons.IsDefaultFolder(ctxFolder)

			Case "folderIcons.Explorer.Button.Help": Return ctxFolder Is Nothing
		End Select
		GetControlVisible = false
	End Function
	
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Public Function GalleryGetItemCount(control As IRibbonControl) As String
		Dim ctxFolder As Folder = Nothing
		Try: ctxFolder = CType(control.Context, Folder)
		Catch ex As System.Exception
		End Try
		If m_foldericons.IsDefaultFolder(ctxFolder) Then Return "0"
		Return CType(m_imagefiles.length, String)
	End Function
	
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Public Function GalleryGetItemLabel(control As IRibbonControl, itemIndex As Integer) As String
		Dim tmp As String
		tmp = m_imagefiles(itemIndex).Substring(m_foldericons.ImageLibraryFolder.Length()+1)
		return tmp.Substring(0, tmp.length()-4)
	End Function
	
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Public Function GalleryGetItemImage(control As IRibbonControl, itemIndex As Integer) As System.Drawing.Bitmap
		Dim bmp As System.Drawing.Bitmap = Nothing
		Try 'icon file can be deleted anytime
			Dim w As Int32 = System.Windows.Forms.SystemInformation.SmallIconSize.Width
			bmp = New System.Drawing.Icon(m_imagefiles(itemIndex), w, w).ToBitmap()
		Catch ex As System.Exception
		End Try
		Return bmp
	End Function

	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Public Sub GalleryItemClick(control As IRibbonControl, itemID As String, itemIndex As Integer)
		Dim ctxFolder As Folder = Nothing
		Try: ctxFolder = CType(control.Context, Folder)	'if called from folder treeview context menu
		Catch ex As System.Exception
		End Try
		m_foldericons.SetIcon(ctxFolder, m_imagefiles(itemIndex))
		m_ribbon.Invalidate
	End Sub
	
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Private Sub OpenLibraryFolder()
		Process.Start(m_foldericons.ImageLibraryFolder)
	End Sub
	
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Private Sub OpenHelpFile()
		System.Diagnostics.Process.Start(AssemblyDirectory & "\Help\index.html")
	End Sub
	
	'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	Public ReadOnly Property AssemblyDirectory() As String
		Get
			Return System.Io.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)
		End Get
	End Property
 
End Class
