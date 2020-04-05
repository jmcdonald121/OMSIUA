VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ODGSMapScan"
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
On Error GoTo Exist_Err

    Dim pGxObj As IGxObject
    Dim pGxFile As IGxFile
    Dim pGxDataset As IGxDataset
    Dim pWsFact As IWorkspaceFactory
    Dim pRWs As IRasterWorkspace
    Dim pRasterDS As IRasterDataset
    
    Set pWsFact = New RasterWorkspaceFactory
    Set pRWs = pWsFact.OpenFromFile(m_pDocPath, 0)
    
    On Error Resume Next
    Set pRasterDS = pRWs.OpenRasterDataset(m_pDocName)
    On Error GoTo Exist_Err
    
    If pRasterDS Is Nothing Then
        Exist = False
    Else
        Exist = True
    End If

Exist_Err:
'    MsgBox Err.Description & vbCrLf & Err.Number
End Property

Public Property Let Name(strDocName As String)
    
    m_pDocName = strDocName
    
End Property

Public Property Get Name() As String
    
    Name = m_pDocName
    
End Property
