Attribute VB_Name = "modPhoto"
Public objPhoto As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)

Public Function CapturePhoto()
    ptrPhoto.FireEvent
    
End Function

Private Function ptrPhoto() As photoClass
    Dim Photo As photoClass
    CopyMemory Photo, objPhoto, 4&
    Set ptrPhoto = Photo
    CopyMemory Photo, 0&, 4&
    
End Function

