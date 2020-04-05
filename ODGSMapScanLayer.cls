VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ODGSMapScanLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_pMxDoc As IMxDocument
Private m_pMap As IMap
Private m_pRLayer As IRasterLayer

Private Sub Class_Initialize()
    
    Set m_pMxDoc = ThisDocument
    Set m_pMap = m_pMxDoc.FocusMap

End Sub

Private Sub Class_Terminate()

End Sub

Public Property Get ODGSLayer(strPath As String, strFileName As String) As ILayer
    
    Dim pRLayer As IRasterLayer
    Dim pLegendInfo As ILegendInfo
    Dim pLegendGroup As ILegendGroup

    Dim strFullPath As String
    
    Dim pODGSMapScan As ODGSMapScan
    
    'Start
    Set pODGSMapScan = New ODGSMapScan
    pODGSMapScan.Path = strPath
    pODGSMapScan.Name = strFileName
    If pODGSMapScan.Exist = True Then
        strFullPath = pODGSMapScan.Path & "\" & pODGSMapScan.Name
        Set pRLayer = New RasterLayer
        pRLayer.CreateFromFilePath (strFullPath)
        
        'Collapse the legend
        Set pLegendInfo = pRLayer
        Set pLegendGroup = pLegendInfo.LegendGroup(0)
        pLegendGroup.Visible = False

        Set m_pRLayer = pRLayer 'May not need this in the future
        Set ODGSLayer = pRLayer
        'Turn off display of the layer in the Table of Contents (TOC)
        ODGSLayer.Visible = False
    Else
        Set ODGSLayer = Nothing
    End If
End Property

