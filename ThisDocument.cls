VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------------------------------------------------------------------
'Title: OMSIUA Appliction
'Date: 20090904
'Version: 1.0
'Abstract: This is an application to load all available information about an
'AUM complaint into an ArcMap. In addition, the application will copy all
'scanned records from the file repository to a temporary directory.
'------------------------------------------------------------------------------
'James McDonald
'GIMS Specialist
'Ohio Division of Geological Survey
'2045 Morse Road
'Columbus, OH  43229-6693
'Ph. (614) 265-6601
'Email: jim.mcdonald@dnr.state.oh.us
'------------------------------------------------------------------------------

Option Explicit

Private m_pMxApp As IMxApplication
Private m_pMxDoc As IMxDocument
Private m_pDisplay As IScreenDisplay
Private m_pPoint As IPoint

Private Sub UIButtonControl1_Click()
'------------------------------------------------------------------------------
'Title: UIButtonControl1_Click
'Date: 20090908
'Version: 1.0
'Abstract: This button resets the ArcMap. It removes all the added layers and
'zooms to full extent.
'------------------------------------------------------------------------------
    Dim pMap As IMap
    Dim pAV As IActiveView
    Dim pGrphcon As IGraphicsContainer
    Dim pEnum As IEnumLayer
    Dim pLayer As ILayer
    Dim pGrpLayer As IGroupLayer
    
    Set m_pMxDoc = ThisDocument
    Set pMap = m_pMxDoc.FocusMap
    Set m_pMxDoc.ActiveView = pMap
    
    'Find the Group layers and remove them from the Table of Contents
    Set pEnum = pMap.Layers
    
    Set pLayer = pEnum.Next
    Do Until pLayer Is Nothing
        If TypeOf pLayer Is IGroupLayer Then
            pMap.DeleteLayer pLayer
        End If
        Set pLayer = pEnum.Next
    Loop
    
    'Remove any graphics elements from the Data Frame. This will not remove graphics elements from the Page Layout
    Set pAV = m_pMxDoc.ActiveView
    Set pGrphcon = pAV
    pGrphcon.DeleteAllElements
    
    m_pMxDoc.ActiveView.Extent = m_pMxDoc.ActiveView.FullExtent
    m_pMxDoc.UpdateContents
    m_pMxDoc.ActiveView.Refresh
    
    Call ExecuteCmd
    
End Sub

Private Function UIButtonControl1_ToolTip() As String
    Dim strMessage As String
    
    strMessage = "Reset and Zoom to Ohio"
    UIButtonControl1_ToolTip = strMessage
    
End Function

Private Sub UIButtonControl2_Click()
    Dim pMSIProject As OMSIUAProject
    
    Set pMSIProject = New OMSIUAProject
    Call pMSIProject.MSIExportPDF

End Sub

Private Function UIButtonControl2_ToolTip() As String
    Dim strMessage As String
    
    strMessage = "Export Vector PDF Files"
    UIButtonControl2_ToolTip = strMessage

End Function

Private Sub UIToolControl1_Select()
    Set m_pMxApp = Application
    Set m_pMxDoc = ThisDocument
    Set m_pDisplay = m_pMxApp.Display

End Sub

Private Sub UIToolControl1_MouseDown(ByVal button As Long, ByVal shift As Long, ByVal x As Long, ByVal y As Long)
'------------------------------------------------------------------------------
'Title: UIButtonControl1_Click
'Date: 20090904
'Version: 1.0
'Abstract: This button zooms to the selected point and adds the map layers.
'------------------------------------------------------------------------------
'    Dim pEnv As IEnvelope
    Dim pMap As IMap
    Dim pSpatialRef As ISpatialReference
    Dim pPointColl As IPointCollection
    Dim pPoly As IPolygon
    Dim pTopoOp As ITopologicalOperator
        
    Dim pPnt1 As IPoint
    Dim pPnt2 As IPoint
    Dim pPnt3 As IPoint
    Dim pPnt4 As IPoint
    
    Dim strStart As String
    Dim strEnd As String
    
    Dim pMSIProject As OMSIUAProject
    
    'START
    strStart = TIME
    
    Set pMap = m_pMxDoc.FocusMap
    Set pSpatialRef = pMap.SpatialReference
    
    Set m_pPoint = m_pDisplay.DisplayTransformation.ToMapPoint(x, y)
    Set m_pPoint.SpatialReference = pSpatialRef
    
    'There are three variations of creating a polygon to find the scanned images
    
    '1. Code for creating a square polygon for finding the scanned images
'    Set pPnt1 = New Point
'    Set pPnt2 = New Point
'    Set pPnt3 = New Point
'    Set pPnt4 = New Point
'
'    pPnt1.PutCoords m_pPoint.x + 2640#, m_pPoint.y + 2640#
'    pPnt2.PutCoords m_pPoint.x + 2640#, m_pPoint.y - 2640#
'    pPnt3.PutCoords m_pPoint.x - 2640#, m_pPoint.y - 2640#
'    pPnt4.PutCoords m_pPoint.x - 2640#, m_pPoint.y + 2640#
'
'    Set pPoly = New Polygon
'    Set pPointColl = New Polygon
'    pPointColl.AddPoint pPnt1
'    pPointColl.AddPoint pPnt2
'    pPointColl.AddPoint pPnt3
'    pPointColl.AddPoint pPnt4
'
'    Set pPoly = pPointColl
'    pPoly.Close
'    Set pTopoOp = pPoly
'    pTopoOp.Simplify 'Needed since creating a polygon using a pointcollection automatically sets the geometry to complex, even if the geometry is non-complex.
'
'    Set pPoly.SpatialReference = pSpatialRef
    
    '2. Code for using a Envelope for finding the scanned images
'    Set pEnv = New Envelope
'    pEnv.PutCoords m_pPoint.x - 2640#, m_pPoint.y - 2640#, m_pPoint.x + 2640#, m_pPoint.y + 2640#
    
    '3. Code for creating and using a buffer for finding the scanned images.
    Set pMSIProject = New OMSIUAProject
    
    Set pTopoOp = m_pPoint
    Set pPoly = pTopoOp.Buffer(2640#)
    
    pMSIProject.ZoomtoAUMLocation m_pPoint
    pMSIProject.LoadLayerFiles
'    pMSIProject.LoadBasemapImages pPoly
    pMSIProject.LoadDRGImages pPoly
    pMSIProject.LoadOSIPImages pPoly
    pMSIProject.Load15MinImages pPoly
'    Call Load24KImages.Load75MinuteImagery("Bedrock Structure", "Z:\IMAGES\MAPS\Bedrock Structure Contours Mylar7.5minute", pPoly)
'    Call Load24KImages.Load75MinuteImagery("Bedrock Geology", "Z:\IMAGES\MAPS\Bedrock Geology Mylar7.5minute", pPoly)
'    Call Load24KImages.Load75MinuteImagery("Bedrock Topography", "Z:\IMAGES\MAPS\Bedrock Topography Mylar7.5minute", pPoly)
    pMSIProject.CopyWellCards pPoly
    pMSIProject.CopyDocuments pPoly
    
    Call ExecuteCmd
    
    strEnd = TIME
    MsgBox "Start = " & strStart & vbCrLf & "End = " & strEnd
    
End Sub

Private Sub ExecuteCmd()
    'URL: http://edndoc.esri.com/arcobjects/9.1/default.asp?URL=/arcobjects/9.1/ArcGISDevHelp/TechnicalDocuments/Guids/ArcMapIds.htm
    Dim pCmdItem As ICommandItem
    'Use ArcID module and the Name of the Save command
    Set pCmdItem = Application.Document.CommandBars.Find(arcid.PanZoom_Pan)
    pCmdItem.Execute
End Sub

Private Function UIToolControl1_ToolTip() As String
    Dim strMessage As String
    
    strMessage = "Zoom to Mine Subsidence Location"
    UIToolControl1_ToolTip = strMessage

End Function

Private Sub UIToolControl2_MouseDown(ByVal button As Long, ByVal shift As Long, ByVal x As Long, ByVal y As Long)
    Dim pMap As IMap
    Dim pSpatialRef As ISpatialReference
    
    Set pMap = m_pMxDoc.FocusMap
    Set pSpatialRef = pMap.SpatialReference
    
    Set m_pPoint = m_pDisplay.DisplayTransformation.ToMapPoint(x, y)
    Set m_pPoint.SpatialReference = pSpatialRef
    
    Call frmMineData_v2.FindMineInfo(m_pPoint)

End Sub

Private Sub UIToolControl2_Select()
    Set m_pMxApp = Application
    Set m_pMxDoc = ThisDocument
    Set m_pDisplay = m_pMxApp.Display

    Call frmMineData_v2.Show(vbModeless)

End Sub

Private Function UIToolControl2_ToolTip() As String
    Dim strMessage As String
    
    strMessage = "Coal Mine Information"
    UIToolControl2_ToolTip = strMessage
    
End Function
