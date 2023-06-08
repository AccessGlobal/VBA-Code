'Código para incluir en un formulario que contenga el objeto InkPicture
'Formulario con dos botones: Grabar y borrar
Option Compare Database
Option Explicit

Private Sub btnBorrar_Click()

    Me.ImgFirma.Ink.DeleteStrokes
    Me.ImgFirma.Requery
    
    Me.ImgMirror.Picture = ""
    
End Sub

Private Sub btnGuardar_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-captura-de-firma/
'                     Destello formativo 337
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Autor             : desconocido
' Adaptado          : Luis Viadel | luisviadel@access-global.net
' Creado            : marzo 2010
' Propósito         : obtener una imagen creada a mano alzada. Por ejemplo, una firma.
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://learn.microsoft.com/en-us/windows/win32/tablet/inkpicture-control-reference
'                     https://learn.microsoft.com/en-us/windows/win32/tablet/inkpicture-control
'                     https://access-global.net/vba-freefile-function/
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim filename As String
Dim filenum As Long
Dim contador As Long
Dim bytArray() As Byte

    filename = Application.CurrentProject.Path & "\firma.GIF"
    filenum = FreeFile()
    
    bytArray = Me.ImgFirma.Ink.Save(IPF_GIF, IPCM_Default)
    
    Open filename For Binary Access Write As filenum
        For contador = 0 To UBound(bytArray)
            Put #filenum, , bytArray(contador)
        Next
    Close #filenum

    Me.ImgMirror.Picture = filename
    
End Sub


