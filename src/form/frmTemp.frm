VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmTemp 
   ClientHeight    =   420
   ClientLeft      =   60
   ClientTop       =   -2550
   ClientWidth     =   1740
   Icon            =   "frmTemp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   420
   ScaleWidth      =   1740
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   375
      ExtentX         =   661
      ExtentY         =   661
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bRunEnd As Boolean
Public sString  As String
Public sType As String

Private Sub Form_Load()
    Dim Fsel As New fileClass
    Dim P As New photoClass
    
    Dim sStr As String

        If sString = "" Then
'Debug.Print 1
'            'sStr = F1.cyDialogSave(0, "保存", "*.JPG|*.JPG|*.BMP|*.BMP")
'            sStr = cyDialogSave(0, "保存", "*.JPG|*.JPG")
'
'Debug.Print 11
'
'            If UCase(Right(sStr, 4)) = ".BMP" Then
'                '文件名不为空则保存
'                If sStr <> "" Then P.cyClipBoardSaveToBmp (sStr)
'
'            Else
'                '文件名不为空则保存
'                If sStr <> "" Then P.cyClipBoardSaveToJpg (sStr)
'
'            End If
'            Unload Me
            
        Else
            
            Web1.Navigate sString
            Do While Not bRunEnd
                DoEvents
            Loop
            bRunEnd = False
            Unload Me
            
        End If

End Sub

Private Sub Web1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
If sType = "Html" Then
    sString = Web1.Document.Body.innerhtml
Else
    sString = Web1.Document.Body.innertext
End If
    bRunEnd = True
End Sub

