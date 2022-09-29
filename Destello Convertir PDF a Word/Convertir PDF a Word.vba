Option Compare Database
Option Explicit

Public Function ConvertPDFToWord(Fichero_Destino As String)
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Fuente            : https://access-global.net/vba-convertir-pdf-a-word
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Título            : CargaFavoritos
' Autor             : Desconocido
' Creación          : desconocida
' Propósito         : convertir un fichero PDF en un fichero Word
' Retorno           : Sin retorno
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Referencias       : https://learn.microsoft.com/es-es/office/vba/api/word.application.changefileopendirectory
'                     https://learn.microsoft.com/es-es/office/vba/api/word.document
'                     https://learn.microsoft.com/es-es/office/vba/api/word.saveas2+
'                     Microsoft Word x.x Object Library
'-----------------------------------------------------------------------------------------------------------------------------------------------
'Test:              : Para adaptar este código en tu aplicación puedes basarte en este procedimiento test. Copiar el bloque siguiente al
'                    portapapeles y pega en el editor de VBA. Descomentar la línea que nos interese y pulsar F5 para ver su funcionamiento.
'
'   Sub ConvertPDFToWord_test()
'
'       Call ConvertPDFToWord (nombre del fichero)
'
'   End Sub
'
'-----------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo LinError
    
    Do While (Fichero_Destino <> "")
   
        ChangeFileOpenDirectory "C:\Users\luisv\Documents" 'Directorio de origen
             
             Documents.Open filename:=Fichero_Destino, ConfirmConversions:=False, ReadOnly:= _
                            False, AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:= _
                            "", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", _
                            Format:=wdOpenFormatAuto, XMLTransform:=""
         
'         ChangeFileOpenDirectory "C:\Users\luisv\Documents" 'Directorio de destino que no incluimos si ambos son iguales
         
             ActiveDocument.SaveAs2 filename:=Replace(Fichero_Destino, ".pdf", ".docx"), FileFormat:=wdFormatXMLDocument _
                                             , LockComments:=False, Password:="", AddToRecentFiles:=True, _
                                             WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
                                             SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
                                             False, CompatibilityMode:=15
         
         ActiveDocument.Close
         Exit Function
    Loop
        
LinError:
    MsgBox "Ha habido un error " & Err.Description
        
End Function
