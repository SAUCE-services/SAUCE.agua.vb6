Attribute VB_Name = "modCollection"
Option Explicit

Public Function collectionExistElement(pCollection As Collection, pIndex As Variant) As Boolean

On Error GoTo existError

    Dim elemento As Object

    Set elemento = pCollection(pIndex)
    
    collectionExistElement = True
    
    Exit Function

existError:
    collectionExistElement = False
    
End Function
