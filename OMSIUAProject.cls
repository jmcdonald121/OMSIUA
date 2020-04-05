VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OMSIUAProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_pMxDoc As IMxDocument
Private m_pMap As IMap

Private Sub Class_Initialize()
    Set m_pMxDoc = ThisDocument
    Set m_pMap = m_pMxDoc.FocusMap

    'Ideas
    '1. Put the higher level dims in this section
    '2. Place the call to the read the document path tables here. Then use a private sub to read the document paths.
    
End Sub

Private Sub Class_Terminate()
        
    Set m_pMap = Nothing
    Set m_pMxDoc = Nothing
    
End Sub

Public Sub ZoomtoAUMLocation(pPoint As IPoint)
    Dim pAUMEnv As IEnvelope
    
    Set pAUMEnv = New Envelope
    pAUMEnv.PutCoords pPoint.x - 2640, pPoint.y - 2640, pPoint.x + 2640, pPoint.y + 2640
    pAUMEnv.Expand 200, 200, False
    
'    Call CreateAORGrahpic(pPoint)
    
    m_pMxDoc.ActiveView.Extent = pAUMEnv
    m_pMxDoc.ActiveView.Refresh

End Sub

Private Sub CreateAORGrahpic(pGeom As IGeometry)
    Dim pPoly As IPolygon
    Dim pPolyline As IGeometryCollection
    Dim pGeomColl As IGeometryCollection
    Dim pSegColl As ISegmentCollection
    Dim pClone As IClone
    Dim pTopoOp As ITopologicalOperator
    Dim pAV As IActiveView
    Dim pGrphcon As IGraphicsContainer
    Dim pCircleElem As IElement
    Dim pFillShpElem As IFillShapeElement
    Dim pFillSym As IFillSymbol
'    Dim pLineElement As IElement
    
    Dim i As Long
    
    Set pTopoOp = pGeom
    Set pPoly = pTopoOp.Buffer(2640#)
    pPoly.SimplifyPreserveFromTo
    
    Set pCircleElem = New CircleElement
    pCircleElem.Geometry = pPoly
    
    Set pFillShpElem = pCircleElem
    pFillShpElem.Symbol = CircleSymbol
    
    Set pAV = m_pMap
    Set pGrphcon = pAV
'    pGrphcon.AddElement pLineElement, 0
    pGrphcon.AddElement pCircleElem, 0

End Sub

Private Function CircleSymbol() As ISimpleFillSymbol
    Dim pSimpleFillSym As IFillSymbol
    Dim pLineSymbol As ISimpleLineSymbol
    Dim pColor As IColor
    Dim pLineColor As IColor

    Set pSimpleFillSym = New SimpleFillSymbol
    Set pLineSymbol = New SimpleLineSymbol
    
    'Set up the fill color
    Set pColor = New RgbColor
    pColor.NullColor = True

    'Set up the outline color, width, and style
    Set pLineColor = New RgbColor
    pLineColor.RGB = vbRed
    pLineSymbol.Color = pLineColor
    pLineSymbol.Width = 2#
    pLineSymbol.Style = esriSLSSolid
    
    'Finish setting up the Polygon symbol, i.e. color and outline
    pSimpleFillSym.Color = pColor
    pSimpleFillSym.Outline = pLineSymbol
    
    Set CircleSymbol = pSimpleFillSym


End Function

Public Sub CopyDocuments(pGeom As IGeometry)
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pFlayer As IFeatureLayer
    Dim pFC As IFeatureClass
    Dim pFCur As IFeatureCursor
    Dim pFeature As IFeature
    
    Dim pStdTableColl As IStandaloneTableCollection
    Dim pStdTable As IStandaloneTable
    Dim pTable As ITable
    Dim pQf As IQueryFilter
    Dim pCursor As ICursor
    Dim pRow As IRow

    Dim pSpatialFilter As ISpatialFilter
    
    Dim lngFldPath As Long
    Dim lngFldFileName As Long
    Dim lngFldNCRDSDocId As Long
    Dim lngFldODGSDocId As Long
    
    Dim i As Integer
    
    Dim strNCRDSDocId As String
    Dim strODGSDocId As String
    Dim strFolder As String
    Dim strFileName As String
    Dim strFullFileName As String
        
    Dim pFSO As FileSystemObject
    Dim pFolder As Folder
    Dim pNewFolder As Folder
    Dim pFiles As Files
    Dim pFile As File
    
    Dim pODGSDocument As ODGSDocument
    
    'Start
    Set pEnumLayer = m_pMap.Layers
    Set pLayer = pEnumLayer.Next
    
    Set pFSO = New FileSystemObject
    
    If (pFSO.FolderExists("D:\OMSIUA\TMP") = False) Then
        Set pNewFolder = pFSO.CreateFolder("D:\OMSIUA\TMP")
    Else
        Set pNewFolder = pFSO.GetFolder("D:\OMSIUA\TMP")
    End If
    
    Do While Not pLayer Is Nothing
        If pLayer.Name = "NCRDS_PTS" Then
            Set pFlayer = pLayer
            Set pFC = pFlayer.FeatureClass
            lngFldNCRDSDocId = pFC.FindField("ODGSDOCID")
            Exit Do
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    
        'Find the Quad Map Scans metadata table
    Set pStdTableColl = m_pMap
    For i = 0 To pStdTableColl.StandaloneTableCount - 1
        Set pStdTable = pStdTableColl.StandaloneTable(i)
        If pStdTable.Name = "ODGSDOCLOCATIONS" Then
            Set pTable = pStdTable
            lngFldODGSDocId = pTable.FindField("ODGSDOCID")
            lngFldPath = pTable.FindField("PATH")
            lngFldFileName = pTable.FindField("FILENAME")
        End If
    Next i
    
    Set pSpatialFilter = New SpatialFilter
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    pSpatialFilter.GeometryField = "SHAPE"
    Set pSpatialFilter.OutputSpatialReference("SHAPE") = pGeom.SpatialReference
    Set pSpatialFilter.Geometry = pGeom
    
    Set pFCur = pFC.Search(pSpatialFilter, False)
    Set pFeature = pFCur.NextFeature
    
    Do While Not pFeature Is Nothing
        strNCRDSDocId = pFeature.value(lngFldNCRDSDocId)
        
        Set pQf = New QueryFilter
        pQf.WhereClause = "[" & pTable.Fields.Field(lngFldODGSDocId).Name & "] = '" & strNCRDSDocId & "'"
        Set pCursor = pTable.Search(pQf, False)
        Set pRow = pCursor.NextRow
        
        Do Until pRow Is Nothing
        
            If VarType(pRow.value(lngFldPath)) = vbNull Then
                strFolder = ""
            Else
                strFolder = pRow.value(lngFldPath)
            End If

            If VarType(pRow.value(lngFldFileName)) = vbNull Then
                strFileName = ""
            Else
                strFileName = pRow.value(lngFldFileName)
            End If
            
            Set pODGSDocument = New ODGSDocument
            pODGSDocument.Path = strFolder
            pODGSDocument.Name = strFileName
            If pODGSDocument.Exist = True Then
                strFullFileName = strFolder & "\" & strFileName
                pFSO.CopyFile strFullFileName, pNewFolder.Path & "\" & strFileName, True
            End If
            Set pRow = pCursor.NextRow
        Loop
        
        Set pFeature = pFCur.NextFeature
    Loop

End Sub

Public Sub CopyWellCards(pGeom As IGeometry)
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pFlayer As IFeatureLayer
    Dim pFC As IFeatureClass
    Dim pFCur As IFeatureCursor
    Dim pFeature As IFeature
    
    Dim pSpatialFilter As ISpatialFilter
    
    Dim lngFldPDFLink As Long
    Dim strFullFileName As String
    Dim strOldFullFileName As String
        
    Dim pFSO As FileSystemObject
    Dim pFolder As Folder
    Dim pNewFolder As Folder
    Dim pFiles As Files
    Dim pFile As File
    
    'Start
    Set pEnumLayer = m_pMap.Layers
    Set pLayer = pEnumLayer.Next
    
    Set pFSO = New FileSystemObject
    
    If (pFSO.FolderExists("D:\OMSIUA\TMP") = False) Then
        Set pNewFolder = pFSO.CreateFolder("D:\OMSIUA\TMP")
    Else
        Set pNewFolder = pFSO.GetFolder("D:\OMSIUA\TMP")
    End If
    
    Do While Not pLayer Is Nothing
        If pLayer.Name = "Oil and Gas Wells" Then
            Set pFlayer = pLayer
            Set pFC = pFlayer.FeatureClass
            lngFldPDFLink = pFC.FindField("PDF_LINK")
            Exit Do
        End If
        Set pLayer = pEnumLayer.Next
    Loop

    Set pSpatialFilter = New SpatialFilter
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    pSpatialFilter.GeometryField = "SHAPE"
    Set pSpatialFilter.OutputSpatialReference("SHAPE") = pGeom.SpatialReference
    Set pSpatialFilter.Geometry = pGeom
    
    Set pFCur = pFC.Search(pSpatialFilter, False)
    Set pFeature = pFCur.NextFeature
    
    Do While Not pFeature Is Nothing
        If StrComp(Left(pFeature.value(lngFldPDFLink), 14), "nowellcard.htm", vbTextCompare) <> 0 Then
            strFullFileName = "Z:\IMAGES\DOCUMENTS\WellCardImages" & "\" & pFeature.value(lngFldPDFLink)
            pFSO.CopyFile strFullFileName, pNewFolder.Path & "\" & Mid(pFeature.value(lngFldPDFLink), 5, 18), True
        End If
        Set pFeature = pFCur.NextFeature
    Loop
    
End Sub


Public Sub Load24kImages(pGeom As IGeometry)
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pNewLayer As ILayer
    Dim pGrpLayer As IGroupLayer
    Dim pQuadLyr As IFeatureLayer
    Dim pQuadFC As IFeatureClass
    Dim pQuadFCur As IFeatureCursor
    Dim pQuadF As IFeature
    
    Dim pSpatialFilter As ISpatialFilter
    
    Dim i As Long
    Dim lngFldQuadName As Long
    Dim lngFldQuadNameScan As Long
    Dim lngFldPath As Long
    Dim lngFldFileName As Long
    Dim strQuadName As String
    
    Dim pStdTableColl As IStandaloneTableCollection
    Dim pStdTable As IStandaloneTable
    Dim pTable As ITable
    
    Dim pQf As IQueryFilter
    Dim pCursor As ICursor
    Dim pRow As IRow
    
    Dim pODGSMapScan As ODGSMapScan
    Dim pODGSMapScanLyr As ODGSMapScanLayer
    
    Dim strFolder As String
    Dim strFileName As String
        
    'START
    Set pEnumLayer = m_pMap.Layers
    
    'Find the Quad layer
    Set pLayer = pEnumLayer.Next
    
    Do While Not pLayer Is Nothing
        If pLayer.Name = "Quad24k" Then
            Set pQuadLyr = pLayer
            Set pQuadFC = pQuadLyr.FeatureClass
            lngFldQuadName = pQuadFC.Fields.FindField("QUADNAME")
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    
    'Find the Quad Map Scans metadata table
    Set pStdTableColl = m_pMap
    For i = 0 To pStdTableColl.StandaloneTableCount - 1
        Set pStdTable = pStdTableColl.StandaloneTable(i)
        If pStdTable.Name = "QUADSCANMAPS" Then
            Set pTable = pStdTable
            lngFldQuadNameScan = pTable.FindField("QUADNAME")
            lngFldPath = pTable.FindField("PATH")
            lngFldFileName = pTable.FindField("FILENAME")
        End If
    Next i
    
    'Create a spatial filter to find the Quad based upon the geometry passed into the proceedure
    Set pSpatialFilter = New SpatialFilter
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    pSpatialFilter.GeometryField = "SHAPE"
    Set pSpatialFilter.OutputSpatialReference("SHAPE") = pGeom.SpatialReference
    Set pSpatialFilter.Geometry = pGeom
    
    Set pQuadFCur = pQuadFC.Search(pSpatialFilter, False)
    
    Set pQuadF = pQuadFCur.NextFeature
    
    Do While Not pQuadF Is Nothing
        strQuadName = pQuadF.value(lngFldQuadName)
        
        'Create new group layer
        Set pGrpLayer = New GroupLayer
        pGrpLayer.Name = "MAP SCANS - " & strQuadName & " 7.5-MINUTE QUADRANGLE"
        pGrpLayer.Expanded = True
        pGrpLayer.Visible = False
        
        'Create a cursor on the Quad metadata table
        Set pQf = New QueryFilter
        pQf.WhereClause = "[" & pTable.Fields.Field(lngFldQuadNameScan).Name & "] = '" & strQuadName & "'"
        Set pCursor = pTable.Search(pQf, False)
        Set pRow = pCursor.NextRow
        
        Do Until pRow Is Nothing
            strFolder = pRow.value(lngFldPath)
            strFileName = pRow.value(lngFldFileName)
            Set pODGSMapScanLyr = New ODGSMapScanLayer
            Set pNewLayer = pODGSMapScanLyr.ODGSLayer(strFolder, strFileName)
            If Not pNewLayer Is Nothing Then
                pGrpLayer.Add pNewLayer
            End If
            Set pRow = pCursor.NextRow
        Loop
        
        m_pMap.AddLayer pGrpLayer  'Add the Group Layer to the currently active dataframe in the MXD.
        m_pMap.MoveLayer pGrpLayer, m_pMap.LayerCount - 1  'Move the Group Layer to the bottom of the Table of Contents.
        
        Set pQuadF = pQuadFCur.NextFeature
        
    Loop


End Sub

Public Sub Load15MinImages(pGeom As IGeometry)
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pNewLayer As ILayer
    Dim pGrpLayer As IGroupLayer
    Dim pQuadLyr As IFeatureLayer
    Dim pQuadFC As IFeatureClass
    Dim pQuadFCur As IFeatureCursor
    Dim pQuadF As IFeature
    
    Dim pSpatialFilter As ISpatialFilter
    
    Dim i As Long
    Dim lngFldQuadName As Long
    Dim lngFldQuadNameScan As Long
    Dim lngFldPath As Long
    Dim lngFldFileName As Long
    Dim strQuadName As String
    
    Dim pStdTableColl As IStandaloneTableCollection
    Dim pStdTable As IStandaloneTable
    Dim pTable As ITable
    
    Dim pQf As IQueryFilter
    Dim pCursor As ICursor
    Dim pRow As IRow
    
    Dim pODGSMapScan As ODGSMapScan
    Dim pODGSMapScanLyr As ODGSMapScanLayer
    
    Dim strFolder As String
    Dim strFileName As String
        
    'START
    Set pEnumLayer = m_pMap.Layers
    
    'Find the Quad layer
    Set pLayer = pEnumLayer.Next
    
    Do While Not pLayer Is Nothing
        If pLayer.Name = "Quad15Minutes" Then
            Set pQuadLyr = pLayer
            Set pQuadFC = pQuadLyr.FeatureClass
            lngFldQuadName = pQuadFC.Fields.FindField("QUADNAME15MIN")
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    
    'Find the Quad Map Scans metadata table
    Set pStdTableColl = m_pMap
    For i = 0 To pStdTableColl.StandaloneTableCount - 1
        Set pStdTable = pStdTableColl.StandaloneTable(i)
'        If pStdTable.Name = "COALMAPS15MIN" Then
        If pStdTable.Name = "QUADMAPS15MIN" Then
            Set pTable = pStdTable
            lngFldQuadNameScan = pTable.FindField("QUADNAME15MIN")
            lngFldPath = pTable.FindField("PATH")
            lngFldFileName = pTable.FindField("FILENAME")
        End If
    Next i
    
    'Create a spatial filter to find the Quad based upon the geometry passed into the proceedure
    Set pSpatialFilter = New SpatialFilter
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    pSpatialFilter.GeometryField = "SHAPE"
    Set pSpatialFilter.OutputSpatialReference("SHAPE") = pGeom.SpatialReference
    Set pSpatialFilter.Geometry = pGeom
    
    Set pQuadFCur = pQuadFC.Search(pSpatialFilter, False)
    
    Set pQuadF = pQuadFCur.NextFeature
    
    Do While Not pQuadF Is Nothing
        strQuadName = pQuadF.value(lngFldQuadName)
        
        'Create new group layer
        Set pGrpLayer = New GroupLayer
        pGrpLayer.Name = "MAP SCANS - " & strQuadName & " 15-MINUTE QUADRANGLE"
        pGrpLayer.Expanded = True
        pGrpLayer.Visible = False
        
        'Create a cursor on the Quad metadata table
        Set pQf = New QueryFilter
        pQf.WhereClause = "[" & pTable.Fields.Field(lngFldQuadNameScan).Name & "] = '" & strQuadName & "'"
        Set pCursor = pTable.Search(pQf, False)
        Set pRow = pCursor.NextRow
        
        Do Until pRow Is Nothing
            strFolder = pRow.value(lngFldPath)
            strFileName = pRow.value(lngFldFileName)
            Set pODGSMapScanLyr = New ODGSMapScanLayer
            Set pNewLayer = pODGSMapScanLyr.ODGSLayer(strFolder, strFileName)
            If Not pNewLayer Is Nothing Then
                pNewLayer.Name = strFolder & "\" & strFileName
                pGrpLayer.Add pNewLayer
            End If
            Set pRow = pCursor.NextRow
        Loop
        
        m_pMap.AddLayer pGrpLayer  'Add the Group Layer to the currently active dataframe in the MXD.
        m_pMap.MoveLayer pGrpLayer, m_pMap.LayerCount - 1  'Move the Group Layer to the bottom of the Table of Contents.
        
        Set pQuadF = pQuadFCur.NextFeature
        
    Loop

End Sub

Public Sub LoadLayerFiles()
    Dim pGxLayer As IGxLayer, pGxNewLayer As IGxLayer, pGxFile As IGxFile
    Dim pLayer As ILayer
    Dim pGrpLayer As IGroupLayer
    
    Dim pStdTableColl As IStandaloneTableCollection
    Dim pStdTable As IStandaloneTable
    Dim pTable As ITable
    Dim pTableSort As ITableSort
    Dim pTblSortLyrs As ITableSort
    Dim pRow As IRow
    Dim pRowLyrs As IRow
    Dim pDataStat As IDataStatistics
    Dim pCursor As ICursor
    Dim pLyrCursor As ICursor
    Dim pQf As IQueryFilter
        
    Dim lngFldLayerName As Long
    Dim lngFldPath As Long
    Dim lngFldGroupOrder As Long
    Dim lngFldGroupName As Long
    Dim lngFldGroupTOCOrder As Long
    Dim lngFldGroupVis As Long
    Dim i As Integer
    
    Dim strLayerName As String
    Dim strPath As String
    Dim lngGroupOrder As Long
    Dim strGroupName As String
    Dim lngGroupTOCOrder As Long
    Dim blnGroupVis As Boolean
        
    Dim pEnumVar As IEnumVariantSimple, value As Variant
    
    Dim pODGSLyr As ODGSLayer
        
    'Start
    'Find the Layer Files metadata table
    Set pStdTableColl = m_pMap
    For i = 0 To pStdTableColl.StandaloneTableCount - 1
        Set pStdTable = pStdTableColl.StandaloneTable(i)
        If pStdTable.Name = "ODGSLAYERS" Then
            Set pTable = pStdTable
            lngFldLayerName = pTable.FindField("LAYERNAME")
            lngFldPath = pTable.FindField("PATH")
            lngFldGroupOrder = pTable.FindField("GROUPORDER")
            lngFldGroupName = pTable.FindField("GROUPNAME")
            lngFldGroupTOCOrder = pTable.FindField("GROUPTOCORDER")
            lngFldGroupVis = pTable.FindField("GROUPVISABLE")
        End If
    Next i
    
    If pStdTable Is Nothing Then
        Exit Sub
    End If
    
    'Sort the Table
    Set pTableSort = New TableSort
    With pTableSort
        .Fields = "GROUPTOCORDER, GROUPNAME"
        .Ascending("GROUPTOCORDER") = True
        .Ascending("GROUPNAME") = True
        Set .QueryFilter = Nothing
        Set .Table = pTable
    End With
    pTableSort.Sort Nothing
    
    Set pCursor = pTableSort.Rows
    
    'Find Unique Values in the Table
    Set pDataStat = New DataStatistics
    pDataStat.Field = "GROUPNAME"
    Set pDataStat.Cursor = pCursor
    
    Set pEnumVar = pDataStat.UniqueValues
    value = pEnumVar.Next
    Do Until IsEmpty(value)
        'Now resort the table based upon the layer order in the group
        Set pQf = New QueryFilter
        pQf.WhereClause = "[GROUPNAME] = '" & value & "'"
        Set pLyrCursor = pTable.Search(pQf, False)
        Set pTblSortLyrs = New TableSort
        With pTblSortLyrs
            .Fields = "GROUPORDER"
            .Ascending("GROUPORDER") = True
            Set .QueryFilter = pQf
            Set .Table = pTable
        End With
        pTblSortLyrs.Sort Nothing
        
        'Get the newly sorted rows and create the new Group and Layers inside the Group
        Set pLyrCursor = pTblSortLyrs.Rows
        Set pRowLyrs = pLyrCursor.NextRow
        
        'Create the new Group
        lngGroupTOCOrder = pRowLyrs.value(lngFldGroupTOCOrder)
        blnGroupVis = pRowLyrs.value(lngFldGroupVis)
        
        Set pGrpLayer = New GroupLayer
        pGrpLayer.Visible = blnGroupVis
        pGrpLayer.Expanded = False
        pGrpLayer.Name = pRowLyrs.value(lngFldGroupName)
        
        'Add layers to the new Group
        Do Until pRowLyrs Is Nothing
            strLayerName = pRowLyrs.value(lngFldLayerName)
            strPath = pRowLyrs.value(lngFldPath)
    
            Set pODGSLyr = New ODGSLayer
            Set pLayer = pODGSLyr.LoadLayer(strPath, strLayerName)
            pGrpLayer.Add pLayer
            Set pRowLyrs = pLyrCursor.NextRow
        Loop
        
        'Add the Group layer to the map
        m_pMap.AddLayer pGrpLayer
        m_pMap.MoveLayer pGrpLayer, lngGroupTOCOrder
        
'        Debug.Print "value - " & value & vbTab & "GroupVis = " & blnGroupVis
        value = pEnumVar.Next
    Loop
    
End Sub

Public Sub LoadBasemapImages(pGeom As IGeometry)
    'In the future, load the paths from the table. Currently, the paths are hard coded into the code.
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pNewLayer As ILayer
    Dim pGrpLayer As IGroupLayer
    Dim pIndexLyr As IFeatureLayer
    Dim pIndexFC As IFeatureClass
    Dim pIndexFCur As IFeatureCursor
    Dim pIndexF As IFeature
    
    Dim pStdTableColl As IStandaloneTableCollection
    Dim pStdTable As IStandaloneTable
    Dim pTable As ITable
    Dim pQf As IQueryFilter
    Dim pCursor As ICursor
    Dim pRow As IRow
    
    Dim pSpatialFilter As ISpatialFilter
    
    Dim i As Long
    Dim lngFldIndexName As Long
    Dim lngFldIndexNameScan As Long
    Dim lngFldPath As Long
    Dim lngFldFileName As Long
    
    Dim strIndexName As String
    Dim strFileName As String
    Dim strFolder As String
    
    Dim pODGSMapScanLyr As ODGSMapScanLayer
        
    'START
    Set pEnumLayer = m_pMap.Layers
    
    'Find the Index layer
    Set pLayer = pEnumLayer.Next
    
    Do While Not pLayer Is Nothing
        If pLayer.Name = "Counties" Then
            Set pIndexLyr = pLayer
            Set pIndexFC = pIndexLyr.FeatureClass
            lngFldIndexName = pIndexFC.Fields.FindField("COUNTY")
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    
    'Create a spatial filter to find the Quad based upon the geometry passed into the proceedure
    Set pSpatialFilter = New SpatialFilter
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    pSpatialFilter.GeometryField = "SHAPE"
    Set pSpatialFilter.OutputSpatialReference("SHAPE") = pGeom.SpatialReference
    Set pSpatialFilter.Geometry = pGeom
    
    Set pIndexFCur = pIndexFC.Search(pSpatialFilter, False)
    
    Set pIndexF = pIndexFCur.NextFeature
    
    'Create new group layer
    Set pGrpLayer = New GroupLayer
    pGrpLayer.Name = "Basemaps"
    pGrpLayer.Expanded = True
    pGrpLayer.Visible = False
    
    'Load the DRGs
    Do While Not pIndexF Is Nothing
        strIndexName = LCase(pIndexF.value(lngFldIndexName))
        
        strFolder = "X:\DATA NAD83\aa Statewide\DRG\State Plane South Counties"
        Set pODGSMapScanLyr = New ODGSMapScanLayer
        Set pNewLayer = pODGSMapScanLyr.ODGSLayer(strFolder, strIndexName)
        If Not pNewLayer Is Nothing Then
            pNewLayer.Name = "DRG - " & UCase(strIndexName) & " COUNTY"
            pGrpLayer.Add pNewLayer
        End If
            
        strFolder = "X:\DATA NAD83\aa Statewide\DRG\State Plane North Counties"
        Set pODGSMapScanLyr = New ODGSMapScanLayer
        Set pNewLayer = pODGSMapScanLyr.ODGSLayer(strFolder, strIndexName)
        If Not pNewLayer Is Nothing Then
            pNewLayer.Name = "DRG - " & UCase(strIndexName) & " COUNTY"
            pGrpLayer.Add pNewLayer
        End If
        
        Set pIndexF = pIndexFCur.NextFeature
    Loop
    
    Set pIndexFCur = pIndexFC.Search(pSpatialFilter, False)
    
    Set pIndexF = pIndexFCur.NextFeature
    
    'Load the OSIP Orthophotos
    Do While Not pIndexF Is Nothing
        strIndexName = LCase(pIndexF.value(lngFldIndexName))
        
        strFolder = "\\Nrgis1\GIS Imagery\OSIP Imagery"
        Set pODGSMapScanLyr = New ODGSMapScanLayer
        Set pNewLayer = pODGSMapScanLyr.ODGSLayer(strFolder, strIndexName & ".sid")
        If Not pNewLayer Is Nothing Then
            pNewLayer.Name = "OSIP IMAGERY - " & UCase(strIndexName) & " COUNTY"
            pGrpLayer.Add pNewLayer
        End If
            
        Set pIndexF = pIndexFCur.NextFeature
    Loop

    m_pMap.AddLayer pGrpLayer  'Add the Group Layer to the currently active dataframe in the MXD.
    m_pMap.MoveLayer pGrpLayer, m_pMap.LayerCount - 1  'Move the Group Layer to the bottom of the Table of Contents.
    
End Sub

Public Sub MSIExportPDF()
    Dim pAV As IActiveView
    Dim pPageLayout As IPageLayout
    
    Dim pEnum As IEnumLayer
    Dim pLayer As ILayer
    Dim pGLayer As IGroupLayer
    
    Dim pFlayer As IFeatureLayer
    Dim pFClass As IFeatureClass

    Dim pGxLayer As IGxLayer, pGxNewLayer As IGxLayer, pGxFile As IGxFile
    Dim pGrpLayer As IGroupLayer
    
    Dim pStdTableColl As IStandaloneTableCollection
    Dim pStdTable As IStandaloneTable
    Dim pTable As ITable
    Dim pTableSort As ITableSort
    Dim pTblSortLyrs As ITableSort
    Dim pRow As IRow
    Dim pRowLyrs As IRow
    Dim pDataStat As IDataStatistics
    Dim pCursor As ICursor
    Dim pLyrCursor As ICursor
    Dim pQf As IQueryFilter
        
    Dim lngFldLayerName As Long
    Dim lngFldPath As Long
    Dim lngFldGroupOrder As Long
    Dim lngFldGroupName As Long
    Dim lngFldGroupTOCOrder As Long
    Dim lngFldGroupVis As Long
    Dim i As Integer
    
    Dim strLayerName As String
    Dim strOMSIUAComplaintNo As String
    
    Dim strPath As String
    Dim lngGroupOrder As Long
    Dim strGroupName As String
    Dim lngGroupTOCOrder As Long
    Dim blnGroupVis As Boolean
        
    Dim pEnumVar As IEnumVariantSimple, value As Variant
    
    Dim pODGSLyr As ODGSLayer
    
    Dim pGrphcon As IGraphicsContainer
    Dim pElem As IElement
    Dim pTxtElem As ITextElement
    
    'START
    'Get the MSI Complaint Number
    strOMSIUAComplaintNo = InputBox("Enter the OMSIUA Claim Number:" & vbCrLf & vbCrLf & "Use the OMSIUA Claim Number and the Owner Last Name" _
        , "OMSIUA Claim Number and Owner Last Name for PDF Files")
    
    'Goto Page layout
    Set pPageLayout = m_pMxDoc.PageLayout
    Set m_pMxDoc.ActiveView = pPageLayout
    
    'Find Basemap Group and DRG layers and turn both of them on
    Set m_pMap = m_pMxDoc.FocusMap
    Set pEnum = m_pMap.Layers
    
    Set pLayer = pEnum.Next
    
    Do Until pLayer Is Nothing
        If pLayer.Name = "Basemaps" Then pLayer.Visible = True
        If pLayer.Name Like "DRG*" Then pLayer.Visible = True
        Set pLayer = pEnum.Next
    Loop

    'Turn off all the group layers
    Set pStdTableColl = m_pMap
    For i = 0 To pStdTableColl.StandaloneTableCount - 1
        Set pStdTable = pStdTableColl.StandaloneTable(i)
        If pStdTable.Name = "ODGSLAYERS" Then
            Set pTable = pStdTable
            lngFldLayerName = pTable.FindField("LAYERNAME")
            lngFldPath = pTable.FindField("PATH")
            lngFldGroupOrder = pTable.FindField("GROUPORDER")
            lngFldGroupName = pTable.FindField("GROUPNAME")
            lngFldGroupTOCOrder = pTable.FindField("GROUPTOCORDER")
            lngFldGroupVis = pTable.FindField("GROUPVISABLE")
        End If
    Next i

    If pStdTable Is Nothing Then
        Exit Sub
    End If

    'Sort the Table
    Set pTableSort = New TableSort
    With pTableSort
        .Fields = "GROUPTOCORDER, GROUPNAME"
        .Ascending("GROUPTOCORDER") = True
        .Ascending("GROUPNAME") = True
        Set .QueryFilter = Nothing
        Set .Table = pTable
    End With
    pTableSort.Sort Nothing

    Set pCursor = pTableSort.Rows

    'Find Unique Values in the Table
    Set pDataStat = New DataStatistics
    pDataStat.Field = "GROUPNAME"
    Set pDataStat.Cursor = pCursor

    Set pEnumVar = pDataStat.UniqueValues
    value = pEnumVar.Next
    
    Do Until IsEmpty(value)
        pEnum.Reset
        Set pLayer = pEnum.Next
        Do Until pLayer Is Nothing
            If pLayer.Name = value Then pLayer.Visible = False
            Set pLayer = pEnum.Next
        Loop
        strLayerName = value
        value = pEnumVar.Next
    Loop
    
    'Now cycle through the group layers, turn them on one by one and export the PDF file
    'First find the Title on the map
    Set pGrphcon = pPageLayout
    pGrphcon.Reset
    Set pElem = pGrphcon.Next
    
    Do Until pElem Is Nothing
        If TypeOf pElem Is ITextElement Then
            Set pTxtElem = pElem
            If pTxtElem.Text = "OMSIUA Layers" Then Exit Do
        End If
        Set pElem = pGrphcon.Next
    Loop
    
    Set pStdTableColl = m_pMap
    For i = 0 To pStdTableColl.StandaloneTableCount - 1
        Set pStdTable = pStdTableColl.StandaloneTable(i)
        If pStdTable.Name = "ODGSLAYERS" Then
            Set pTable = pStdTable
            lngFldLayerName = pTable.FindField("LAYERNAME")
            lngFldPath = pTable.FindField("PATH")
            lngFldGroupOrder = pTable.FindField("GROUPORDER")
            lngFldGroupName = pTable.FindField("GROUPNAME")
            lngFldGroupTOCOrder = pTable.FindField("GROUPTOCORDER")
            lngFldGroupVis = pTable.FindField("GROUPVISABLE")
        End If
    Next i

    If pStdTable Is Nothing Then
        Exit Sub
    End If

    'Sort the Table
    Set pTableSort = New TableSort
    With pTableSort
        .Fields = "GROUPTOCORDER, GROUPNAME"
        .Ascending("GROUPTOCORDER") = True
        .Ascending("GROUPNAME") = True
        Set .QueryFilter = Nothing
        Set .Table = pTable
    End With
    pTableSort.Sort Nothing

    Set pCursor = pTableSort.Rows

    'Find Unique Values in the Table
    Set pDataStat = New DataStatistics
    pDataStat.Field = "GROUPNAME"
    Set pDataStat.Cursor = pCursor
    Set pEnumVar = pDataStat.UniqueValues
    
    value = pEnumVar.Next
    Do Until IsEmpty(value)
        pEnum.Reset
        Set pLayer = pEnum.Next
        Do Until pLayer Is Nothing
            If pLayer.Name = value Then
                pLayer.Visible = True
                strLayerName = value
                pTxtElem.Text = strLayerName
                Call ExportActiveView.ExportActiveView(strOMSIUAComplaintNo & "_" & strLayerName)
                pLayer.Visible = False
            End If
            Set pLayer = pEnum.Next
        Loop
        value = pEnumVar.Next
    Loop
    
    pTxtElem.Text = "OMSIUA Layers"
    
    Set m_pMxDoc.ActiveView = m_pMap
    
    m_pMxDoc.UpdateContents
    m_pMxDoc.ActiveView.Refresh
    
End Sub

Public Sub LoadDRGImages(pGeom As IGeometry)
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pNewLayer As ILayer
    Dim pGrpLayer As IGroupLayer
    Dim pQuadLyr As IFeatureLayer
    Dim pQuadFC As IFeatureClass
    Dim pQuadFCur As IFeatureCursor
    Dim pQuadF As IFeature
    
    Dim pSpatialFilter As ISpatialFilter
    
    Dim i As Long
    Dim lngFldQuadName As Long
    Dim lngFldQuadNameScan As Long
    Dim lngFldPath As Long
    Dim lngFldFileName As Long
    Dim strQuadName As String
    
    Dim pStdTableColl As IStandaloneTableCollection
    Dim pStdTable As IStandaloneTable
    Dim pTable As ITable
    
    Dim pQf As IQueryFilter
    Dim pCursor As ICursor
    Dim pRow As IRow
    
    Dim pODGSMapScan As ODGSMapScan
    Dim pODGSMapScanLyr As ODGSMapScanLayer
    
    Dim strFolder As String
    Dim strFileName As String
        
    'START
    Set pEnumLayer = m_pMap.Layers
    
    'Find the Quad layer
    Set pLayer = pEnumLayer.Next
    
    Do While Not pLayer Is Nothing
        If pLayer.Name = "Counties" Then
            Set pQuadLyr = pLayer
            Set pQuadFC = pQuadLyr.FeatureClass
            lngFldQuadName = pQuadFC.Fields.FindField("COUNTY")
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    
    pEnumLayer.Reset
    'Find the Quad layer
    Set pLayer = pEnumLayer.Next

    Do While Not pLayer Is Nothing
        If pLayer.Name = "Basemaps" Then
            Set pGrpLayer = pLayer
            Exit Do
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    
    If pLayer Is Nothing Then
        Set pGrpLayer = New GroupLayer
        pGrpLayer.Name = "Basemaps"
        pGrpLayer.Expanded = True
        pGrpLayer.Visible = False
    End If
    
    
    'Find the Quad Map Scans metadata table
    Set pStdTableColl = m_pMap
    For i = 0 To pStdTableColl.StandaloneTableCount - 1
        Set pStdTable = pStdTableColl.StandaloneTable(i)
        If pStdTable.Name = "DRGCOUNTIES" Then
            Set pTable = pStdTable
            lngFldQuadNameScan = pTable.FindField("COUNTY")
            lngFldPath = pTable.FindField("PATH")
            lngFldFileName = pTable.FindField("FILENAME")
        End If
    Next i
    
    'Create a spatial filter to find the Quad based upon the geometry passed into the proceedure
    Set pSpatialFilter = New SpatialFilter
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    pSpatialFilter.GeometryField = "SHAPE"
    Set pSpatialFilter.OutputSpatialReference("SHAPE") = pGeom.SpatialReference
    Set pSpatialFilter.Geometry = pGeom
    
    Set pQuadFCur = pQuadFC.Search(pSpatialFilter, False)
    
    Set pQuadF = pQuadFCur.NextFeature
    
    Do While Not pQuadF Is Nothing
        strQuadName = pQuadF.value(lngFldQuadName)
        
        
        'Create a cursor on the Quad metadata table
        Set pQf = New QueryFilter
        pQf.WhereClause = "[" & pTable.Fields.Field(lngFldQuadNameScan).Name & "] = '" & strQuadName & "'"
        Set pCursor = pTable.Search(pQf, False)
        Set pRow = pCursor.NextRow
        
        Do Until pRow Is Nothing
            strFolder = pRow.value(lngFldPath)
            strFileName = pRow.value(lngFldFileName)
            Set pODGSMapScanLyr = New ODGSMapScanLayer
            Set pNewLayer = pODGSMapScanLyr.ODGSLayer(strFolder, strFileName)
            If Not pNewLayer Is Nothing Then
                pNewLayer.Name = "DRG - " & UCase(strQuadName) & " COUNTY"
                pGrpLayer.Add pNewLayer
            End If
            Set pRow = pCursor.NextRow
        Loop
        
        Set pQuadF = pQuadFCur.NextFeature
        
    Loop
    
    pEnumLayer.Reset
    'Find the Quad layer
    Set pLayer = pEnumLayer.Next

    Do While Not pLayer Is Nothing
        If pLayer.Name = "Basemaps" Then
            Exit Do
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    
    If pLayer Is Nothing Then
        m_pMap.AddLayer pGrpLayer  'Add the Group Layer to the currently active dataframe in the MXD.
        m_pMap.MoveLayer pGrpLayer, m_pMap.LayerCount - 1  'Move the Group Layer to the bottom of the Table of Contents.
    End If

End Sub

Public Sub LoadOSIPImages(pGeom As IGeometry)
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pNewLayer As ILayer
    Dim pGrpLayer As IGroupLayer
    Dim pQuadLyr As IFeatureLayer
    Dim pQuadFC As IFeatureClass
    Dim pQuadFCur As IFeatureCursor
    Dim pQuadF As IFeature
    
    Dim pSpatialFilter As ISpatialFilter
    
    Dim i As Long
    Dim lngFldQuadName As Long
    Dim lngFldQuadNameScan As Long
    Dim lngFldPath As Long
    Dim lngFldFileName As Long
    Dim strQuadName As String
    
    Dim pStdTableColl As IStandaloneTableCollection
    Dim pStdTable As IStandaloneTable
    Dim pTable As ITable
    
    Dim pQf As IQueryFilter
    Dim pCursor As ICursor
    Dim pRow As IRow
    
    Dim pODGSMapScan As ODGSMapScan
    Dim pODGSMapScanLyr As ODGSMapScanLayer
    
    Dim strFolder As String
    Dim strFileName As String
        
    'START
    Set pEnumLayer = m_pMap.Layers
    
    'Find the Quad layer
    Set pLayer = pEnumLayer.Next
    
    Do While Not pLayer Is Nothing
        If pLayer.Name = "Counties" Then
            Set pQuadLyr = pLayer
            Set pQuadFC = pQuadLyr.FeatureClass
            lngFldQuadName = pQuadFC.Fields.FindField("COUNTY")
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    
    Debug.Print pQuadFC.AliasName
    
'    'START again
'    Set pEnumLayer = m_pMap.Layers

    pEnumLayer.Reset
    'Find the Quad layer
    Set pLayer = pEnumLayer.Next

    Do While Not pLayer Is Nothing
        If pLayer.Name = "Basemaps" Then
            Set pGrpLayer = pLayer
            Exit Do
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    
    Debug.Print
    Debug.Print pQuadFC.AliasName
    
    
    If pLayer Is Nothing Then
        Set pGrpLayer = New GroupLayer
        pGrpLayer.Name = "Basemaps"
        pGrpLayer.Expanded = True
        pGrpLayer.Visible = False
    End If
    
    
    'Find the Quad Map Scans metadata table
    Set pStdTableColl = m_pMap
    For i = 0 To pStdTableColl.StandaloneTableCount - 1
        Set pStdTable = pStdTableColl.StandaloneTable(i)
        If pStdTable.Name = "OSIPIMAGERY" Then
            Set pTable = pStdTable
            lngFldQuadNameScan = pTable.FindField("COUNTY")
            lngFldPath = pTable.FindField("PATH")
            lngFldFileName = pTable.FindField("FILENAME")
        End If
    Next i
    
    'Create a spatial filter to find the Quad based upon the geometry passed into the proceedure
    Set pSpatialFilter = New SpatialFilter
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    pSpatialFilter.GeometryField = "SHAPE"
    Set pSpatialFilter.OutputSpatialReference("SHAPE") = pGeom.SpatialReference
    Set pSpatialFilter.Geometry = pGeom
    
    Set pQuadFCur = pQuadFC.Search(pSpatialFilter, False)
    
    Set pQuadF = pQuadFCur.NextFeature
    
    Do While Not pQuadF Is Nothing
        strQuadName = pQuadF.value(lngFldQuadName)
        
        
        'Create a cursor on the Quad metadata table
        Set pQf = New QueryFilter
        pQf.WhereClause = "[" & pTable.Fields.Field(lngFldQuadNameScan).Name & "] = '" & strQuadName & "'"
        Set pCursor = pTable.Search(pQf, False)
        Set pRow = pCursor.NextRow
        
        Do Until pRow Is Nothing
            strFolder = pRow.value(lngFldPath)
            strFileName = pRow.value(lngFldFileName)
            Set pODGSMapScanLyr = New ODGSMapScanLayer
            Set pNewLayer = pODGSMapScanLyr.ODGSLayer(strFolder, strFileName)
            If Not pNewLayer Is Nothing Then
                pNewLayer.Name = "OSIP IMAGERY - " & UCase(strQuadName) & " COUNTY"
                pGrpLayer.Add pNewLayer
            End If
            Set pRow = pCursor.NextRow
        Loop
        
        Set pQuadF = pQuadFCur.NextFeature
        
    Loop
    
    pEnumLayer.Reset
    'Find the Quad layer
    Set pLayer = pEnumLayer.Next

    Do While Not pLayer Is Nothing
        If pLayer.Name = "Basemaps" Then
            Exit Do
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    
    If pLayer Is Nothing Then
        m_pMap.AddLayer pGrpLayer  'Add the Group Layer to the currently active dataframe in the MXD.
        m_pMap.MoveLayer pGrpLayer, m_pMap.LayerCount - 1  'Move the Group Layer to the bottom of the Table of Contents.
    End If

End Sub

