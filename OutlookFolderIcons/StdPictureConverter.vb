Public Class StdPictureConverter
	Inherits System.Windows.Forms.AxHost
	
	Public Sub New()
		MyBase.New("59EE46BA-677D-4d20-BF10-8D8067CB8B33")
	End Sub
	
	Public Shared Function ImageToIPicture(img As System.Drawing.Image) As Microsoft.Interop.Stdole.StdPicture
		ImageToIPicture = CType(System.Windows.Forms.AxHost.GetIPictureFromPicture(img), Microsoft.Interop.Stdole.StdPicture)
	End Function
	
	Public Shared Function IPictureToImage(img As Microsoft.Interop.Stdole.StdPicture) As System.Drawing.Image
		IPictureToImage = System.Windows.Forms.AxHost.GetPictureFromIPicture(img)
	End Function
	
	Public Shared Function LoadIPictureDisp(imagePath As String) As Microsoft.Interop.Stdole.StdPicture
		Try:
			Dim bmp As New system.Drawing.bitmap(imagePath)
			bmp.MakeTransparent(System.Drawing.Color.FromArgb(0, 255, 0, 255)) 'bmp.GetPixel(0, 0))
			Return ImageToIPicture(bmp)
		Catch ex As System.Exception
			'trace.TraceError(ex.ToString())
			return nothing
		End Try
	End Function
	
	Public Shared Function ImageToIPictureDisp(img As system.Drawing.Image) As Microsoft.Interop.Stdole.IPictureDisp
		ImageToIPictureDisp = CType(System.Windows.Forms.AxHost.GetIPictureDispFromPicture(img), Microsoft.Interop.Stdole.IPictureDisp)
	End Function
	
End Class
