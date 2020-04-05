VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ODGSLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub

Public Property Get LoadLayer(strFolder As String, strFileName As String) As ILayer
    Dim pGxLayer As IGxLayer, pGxNewLayer As IGxLayer, pGxFile As IGxFile
    
    Set pGxLayer = New esriCatalog.GxLayer
    Set pGxFile = pGxLayer
    pGxFile.Path = strFolder & "\" & strFileName
    pGxFile.Open
    Set LoadLayer = pGxLayer.Layer

End Property



