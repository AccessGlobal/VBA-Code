'Módulo estándar: modImagenes
Option Compare Database
Option Explicit

Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub Guardar_Galeria(ByVal Path_Imagen As String, ByVal idgallery As Integer, ByVal ext As String, ByVal nom As String, ByVal M As String, ByVal nombrefichero As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/access-imagemagick/
'                   . Destello formativo 384
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : Guardar_Galeria
' Autor original    : Luis Viadel | luisviadel@access-global.net
' Creado            : 2010
' Propósito         : guardar una galería de imágenes manipuladas mediante el complemento ImageMagick (contiene BBDD ejemplo)
' Retorno           : Dirección de la carpeta seleccionada
' Información       : https://imagemagick.org/
' Complemento       : es necesario descargar e instalar en el equipo el complemento desde la web anterior:
'                     https://imagemagick.org/script/download.php
' Argumento/s       : La sintaxis de la función consta del siguientes argumento:
'                     Parte           Modo              Descripción
'-----------------------------------------------------------------------------------------------------------------------------------------------
'                     Path_Imagen     Obligatorio       Dirección de la carpeta que contiene las imágenes
'                     idgallery       Obligatorio       id de la galería donde guardar las imágenes
'                     ext             Obligatorio       extendión de los ficheros que se subirán
'                     idgallery       Obligatorio       Dirección en la que comenzará la búsqueda
'                     nom             Obligatorio       nombre del fichero de destino
'                     M               Obligatorio       número de fichero
'                     nom             Obligatorio       nombre del fichero
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim ByteImage() As Byte
Dim datos As String, watermark As String
Dim objImg As Object
Dim FicheroDestino As String
Dim orden As Integer
Dim rstTable As DAO.Recordset

    watermark = CurrentProject.Path & "\Imágenes\Watermark.png"
    
    FicheroDestino = CurrentProject.Path & "\temp\ficherotemporal.jpg"

'Transforma la imagen en el formato y características adecuadas
    Set objImg = CreateObject("ImageMagickObject.MagickImage.1")
        objImg.Convert Path_Imagen, "-format", "jpg", "-resize", "1024x780", "-density", "72", FicheroDestino
        objImg.Composite "-dissolve", "40%", "-gravity", "center", watermark, FicheroDestino, nom
        Kill FicheroDestino
    Set objImg = Nothing

    Open nom For Binary Access Read As #1

    ReDim ByteImage(1 To LOF(1))
        Get #1, , ByteImage
    Close #1

    datos = encodeBase64(ByteImage)
    orden = Int(M)

    Set rstTable = CurrentDb.OpenRecordset("galimg")
        rstTable.AddNew
            rstTable!idgallery = idgallery
            rstTable!galimgtxt = datos
            rstTable!galimgext = ext
            rstTable!galimgnom = left(nombrefichero, Len(nombrefichero) - 4)
            rstTable!galimgfa = Format(Date, "Short date")
            rstTable!galimgor = orden
        rstTable.Update
    rstTable.Close
    Set rstTable = Nothing
 
End Sub

'módulo estándar del formulario "Galería"

Private Sub NuevaGaleria_Click()
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-busca-carpeta/
'                   . Destello formativo 383
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : NuevaGaleria_Click
' Autor original    : Luis Viadel | luisviadel@access-global.net
' Creado            : 15/08/2015
' Propósito         : Carga las miniaturas en un formulario y graba todas las imágenes en la base de
'                     datos que se encuentren en una carpeta seleccionada con el formato, tamaño y
'                     resolución estipulados (JPG, 1024x768 y 72 ppp)
'-----------------------------------------------------------------------------------------------------------------------------------------------
Dim ctrl As control
Dim I As Integer, N As Integer, tot As Integer
Dim Titulo As String, codprodu As String
Dim Path_inicial As String, carpeta As String, Nombre As String, ext As String
Dim M As String
Dim WshShell As Object
Dim fs, fol, F
Dim Fichero_Destino As String, Fichero_Destino_mini As String, nombrefichero As String
Dim objImg As Object
Dim idprodu As Integer, idgallery As Integer
Dim rstTable As DAO.Recordset
   
    idprodu = 66
    codprodu = DLookup("[produtcod]", "productos", "[idprodut]=" & idprodu)

'Seleccionamos la carpeta donde se encuentran las imágenes
    Set WshShell = CreateObject("WScript.Shell")
    
    Titulo = "Seleccione la carpeta donde se encuentran las imágenes"
    Path_inicial = WshShell.ExpandEnvironmentStrings("%USERDOMAIN%")
    
    carpeta = Busca_Carpeta(Titulo, Path_inicial)
    
    If carpeta = "" Then Exit Sub

'Recorremos todos los ficheros de la carpeta seleccionada
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set fol = fs.GetFolder(carpeta)
    
        tot = fol.Files.Count
        
        If tot = 0 Then Exit Sub
        
        If tot > 5 Then
            MsgBox "La galería puede contener un máximo de 5 imágenes", vbCritical + vbOKOnly, "Error"
            Exit Sub
        End If
          
        I = 1
        N = 1 'Contador de imágenes

'Graba la nueva galería en la tabla de producto
        Set rstTable = CurrentDb.OpenRecordset("SELECT* FROM productos WHERE idprodut=" & idprodu)
            rstTable.Edit
                rstTable!produgal = -1
            rstTable.Update
            rstTable.Close
        Set rstTable = Nothing

'Creamos la galería y recuperamos su idgallery
        Set rstTable = CurrentDb.OpenRecordset("gallery")
            rstTable.AddNew
                rstTable!idprodu = idprodu
                rstTable!gallerynom = "Producto " & codprodu
            rstTable.Update
            rstTable.Close
        Set rstTable = Nothing
        
        idgallery = DMax("[idgallery]", "gallery")
        
        For Each F In fol.Files
            If Len(N) = 2 Then
                M = "0" & N
            Else
                M = N
            End If
            ext = LCase(Right(F.Name, 3))
                If ext = "bmp" Or ext = "jpg" Or ext = "png" Then
                    Nombre = carpeta & "\" & F.Name
'Guardamos el fichero en la BD
                    nombrefichero = codprodu & "_" & M & "." & ext
                    Fichero_Destino = carpeta & "\" & nombrefichero
                    Fichero_Destino_mini = carpeta & "\" & codprodu & "_" & M & "_mini." & ext
                    Call Guardar_Galeria(Nombre, idgallery, ext, Fichero_Destino, M, nombrefichero)
                
'Colocamos la miniatura en la ficha del producto
                    Set objImg = CreateObject("ImageMagickObject.MagickImage.1")
                        objImg.Convert Fichero_Destino, "-strip", "-thumbnail", "150x150", "-unsharp", "0x.5", Fichero_Destino_mini
                    Set objImg = Nothing
                                
                    For Each ctrl In Form_Galeria.Controls
                        If ctrl.Name = "Imagen" & I Then
                            ctrl.Picture = Fichero_Destino_mini
                            ctrl.Visible = True
                        End If
                        If ctrl.Name = "ImgTxt" & I Then
                            ctrl = codprodu & "_" & M & "." & ext
                            ctrl.Visible = True
                        End If
                        If ctrl.Name = "Ver" & I Then
                            ctrl.Visible = True
                        End If
                        If ctrl.Name = "btn" & I Then
                            ctrl.Visible = True
                        End If
                        If ctrl.Name = "Etq" & I Then
                            ctrl.Visible = True
                        End If
                    Next ctrl
                    I = I + 1
                    N = N + 1
        
'Borra la imagen que ha modificado y guardado en la base de datos como imagen de galería
                    Kill Fichero_Destino
                    Kill Fichero_Destino_mini
                End If
        Next F
    
    Set fs = Nothing
    Set fol = Nothing
        
End Sub
