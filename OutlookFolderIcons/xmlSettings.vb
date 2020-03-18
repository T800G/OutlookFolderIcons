Imports System.Xml
Imports Microsoft.Office.Interop
Imports System.Windows.Forms
imports system.Security.Cryptography
Imports System.Text

Public Class xmlSettings
	Private Const m_defaultXML As String = "<?xml version='1.0'?><!-- do not edit manually --><OUTLOOK></OUTLOOK>"
	Private m_xmlDoc As Xml.XmlDocument
	Private m_savefolder As String
	
	Public ReadOnly Property SettingsFolder() As String
		Get
			return m_savefolder
		End Get
	End Property
	
	Public ReadOnly Property ImageLibraryFolder() As String
		Get
			return (m_savefolder & "\Library")
		End Get
	End Property
	
	
	Public Sub New()
	    m_xmlDoc = New Xml.XmlDocument
		m_xmlDoc.LoadXml(m_defaultXML)
		m_savefolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\Outlook Folder Icons"
	    Dim fpath = m_savefolder & "\settings.xml"
	    If system.IO.File.Exists(fpath) Then
	    	Try: m_xmlDoc.Load(fpath)
	    	Catch ex As System.Exception
				MessageBox.Show(("Error loading settings file" & vbCrLf & fpath), "Folder Icons", MessageBoxButtons.OK, MessageBoxIcon.Warning)
				'trace.TraceError(("Error loading settings file" & fpath))
				m_xmlDoc.LoadXml(m_defaultXML)
				Exit sub
	    	End Try
	    Else
	    	m_xmlDoc.LoadXml(m_defaultXML)
	    	Save()
	    End If
	End Sub

	Public Sub Save()
		If Not system.IO.Directory.Exists(m_savefolder) Then
			Try: system.IO.Directory.CreateDirectory(m_savefolder)
			Catch ex As System.Exception
				MessageBox.Show(("Error creating folder" & vbcrlf & m_savefolder),"Folder Icons", MessageBoxButtons.OK, MessageBoxIcon.Warning)
				'trace.TraceError(("Error creating folder" & m_savefolder))
				Exit Sub
	    	End Try
		End If
		Try: m_xmlDoc.Save(m_savefolder & "\settings.xml")
		Catch ex As System.Exception
			MessageBox.Show(("Error saving settings to" & vbcrlf & m_savefolder & "\settings.xml"),"Folder Icons", MessageBoxButtons.OK, MessageBoxIcon.Warning)
			'trace.TraceError(("Error saving settings to" & m_savefolder & "\settings.xml"))
    	End Try
	End Sub
	
	Public Function GetIconPath(storeID As String, entryID As String) As String
	    Dim pNode As XML.XmlNode
	    pNode = m_xmlDoc.SelectSingleNode("//OUTLOOK/STORE[@storeidMD5='" & GetMD5Hash(storeID) & "']/FOLDER[@entryID='" & entryID & "']")
	    If pNode Is Nothing Then Return vbnullstring
	    GetIconPath = pNode.InnerText
	End Function

	Public Sub SetIconPath(storeID As String, entryID As String, iconPath As String)
		If m_xmlDoc Is Nothing Then Exit Sub
	    Dim pRoot As xml.XmlNode = m_xmlDoc.SelectSingleNode("//OUTLOOK")
	    If pRoot Is Nothing Then
			'trace.TraceError("critical error: no OUTLOOK node") 'not my xml?
	        Exit Sub
	    End If
	    Dim md5 As String  = GetMD5Hash(storeID)
	    Dim pStore As xml.XmlNode = pRoot.SelectSingleNode("//OUTLOOK/STORE[@storeidMD5='" & md5 & "']")
	    If pStore Is Nothing Then
	    	pStore = AddXMLNode(pRoot, "STORE", vbNullString, vbNullString, "storeidMD5", md5)
	        If pStore Is Nothing Then
	            'trace.TraceError("AddXMLNode(STORE) failed")
	            Exit Sub
	        End If
	    End If
	    Dim pFolder As xml.XmlNode = pStore.SelectSingleNode("//STORE/FOLDER[@entryID='" & entryID & "']")
	    If pFolder Is Nothing Then
	        pFolder = AddXMLNode(pStore, "FOLDER", vbNullString, vbNullString, "entryID", entryID)
	        If pFolder Is Nothing Then
	            'trace.TraceError("AddXMLNode(FOLDER) failed")
	            Exit Sub
	        End If
	    End If
	    pFolder.innertext = iconPath
		Save()
	End Sub

	Public Sub DeleteSetting(storeID As String, entryID As String)
	    If m_xmlDoc Is Nothing Then Exit Sub
	    Dim pNode As xml.XmlNode = m_xmlDoc.SelectSingleNode("//OUTLOOK/STORE[@storeidMD5='" & GetMD5Hash(storeID) & "']/FOLDER[@entryID='" & entryID & "']")
	    If pNode Is Nothing Then Exit Sub
	    Dim parent As xml.XmlNode = pNode.parentNode
	    parent.RemoveChild(pNode)
	    If parent.ChildNodes.Count = 0 Then parent.parentNode.RemoveChild(parent) 'no child nodes left, remove self (store level)
		Save()
	End Sub
	
	Private Function AddXMLNode(ByRef parentNode As xml.XmlNode, ByVal nodeName As String, _
	 							ByVal nodeText As String, ByVal namespaceURI As String,  _
	 							ByVal attributeName As String, ByVal attributeValue As String) As xml.XmlNode
		Try:AddXMLNode = parentNode.OwnerDocument.createNode(xml.XmlNodeType.Element, nodeName, namespaceURI)
		    Dim attr As xml.XmlAttribute = parentNode.OwnerDocument.createAttribute(attributeName)
		    attr.Value = attributeValue
		    AddXMLNode.Attributes.setNamedItem(attr)
		    AddXMLNode.innerText = nodeText
		    parentNode.appendChild(AddXMLNode)
	    Catch ex As System.Exception
	    	'trace.TraceError(ex.ToString())
	    	AddXMLNode = nothing
		End Try
	End Function
	
	Function GetMD5Hash(ByVal input As String) As String
		Using md5Hash As MD5 = MD5.Create()
		Dim data As Byte() = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(input))
		Dim sBuilder As New StringBuilder()
		Dim i As Integer
		For i = 0 To data.Length - 1
		    sBuilder.Append(data(i).ToString("x2"))
		Next i
		Return sBuilder.ToString()
		End Using
	End Function
	
End Class
