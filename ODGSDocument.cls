VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ODGSDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_pDocExists As Boolean
Private m_pDocPath As String
Private m_pDocName As String

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub

Public Property Let Path(strDocPath As String)

    m_pDocPath = strDocPath
    
End Property

Public Property Get Path() As String

    Path = m_pDocPath

End Property

Public Property Let Exist(blnExist As Boolean)
    
    m_pDocExists = False
    
End Property

Public Property Get Exist() As Boolean
    Dim pFSO As FileSystemObject
    Dim pFolder As Folder
    Dim pFiles As Files
    Dim pFile As File
    
    Set pFSO = New FileSystemObject
    
    Exist = pFSO.FileExists(m_pDocPath & "\" & m_pDocName)
    
End Property

Public Property Let Name(strDocName As String)
    
    m_pDocName = strDocName
    
End Property

Public Property Get Name() As String
    
    Name = m_pDocName
    
End Property

