VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMineData_v2 
   Caption         =   "OMSIUA AUM Information"
   ClientHeight    =   11025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14490
   OleObjectBlob   =   "frmMineData_v2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMineData_v2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Note: Calls ODGSRelClass

Option Explicit
Private m_pMxDoc As IMxDocument
Private m_pMap As IMap
Private m_pLayer As ILayer
Private m_pFLayer As IFeatureLayer
Private m_pFC As IFeatureClass
Private m_strMineAPI As String

Private Sub UserForm_Initialize()
    Dim pEnumLayer As IEnumLayer
    
    Set m_pMxDoc = ThisDocument
    Set m_pMap = m_pMxDoc.FocusMap
    
    Set pEnumLayer = m_pMap.Layers
    
    Set m_pLayer = pEnumLayer.Next
    Do Until m_pLayer Is Nothing
        If m_pLayer.Name = "Underground Mine" Then
            Set m_pFLayer = m_pLayer
        End If
        Set m_pLayer = pEnumLayer.Next
    Loop
    
End Sub

Private Sub UserForm_Terminate()
    Set m_pMxDoc = Nothing
    Set m_pMap = Nothing
    Set m_pLayer = Nothing
    Set m_pFLayer = Nothing
    Set m_pFC = Nothing
    
    Call ExecuteCmd
End Sub

Private Sub ExecuteCmd()
    'URL: http://edndoc.esri.com/arcobjects/9.1/default.asp?URL=/arcobjects/9.1/ArcGISDevHelp/TechnicalDocuments/Guids/ArcMapIds.htm
    Dim pCmdItem As ICommandItem
    'Use ArcID module and the Name of the Save command
    Set pCmdItem = Application.Document.CommandBars.Find(arcid.PanZoom_Pan)
    pCmdItem.Execute
End Sub

Public Sub FindMineInfo(pGeom As IGeometry)
    Dim pFSel As IFeatureSelection
    Dim pFC As IFeatureClass
    Dim lngCount As Long
    
    Dim pQf As IQueryFilter

    Dim pSpatialFilter As ISpatialFilter
    Dim pFCur As IFeatureCursor
    Dim pF As IFeature
    Dim lngFldMineAPI As Long
    
    Set pFSel = m_pFLayer
    
    Set pSpatialFilter = New SpatialFilter
    pSpatialFilter.SpatialRel = esriSpatialRelIntersects
    pSpatialFilter.GeometryField = "SHAPE"
    Set pSpatialFilter.OutputSpatialReference("SHAPE") = pGeom.SpatialReference
    Set pSpatialFilter.Geometry = pGeom
    
    Set m_pFC = m_pFLayer.FeatureClass
    lngCount = m_pFC.FeatureCount(pSpatialFilter)
    
    If lngCount >= 1 Then
        Set pFCur = m_pFLayer.Search(pSpatialFilter, False)
        
        Set pF = pFCur.NextFeature
        lngFldMineAPI = pF.Fields.FindField("MINE_API")
        
        m_strMineAPI = pF.value(lngFldMineAPI)
        
        Set pQf = New QueryFilter
        pQf.WhereClause = "[MINE_API] = '" & m_strMineAPI & "'"
        
        m_pMxDoc.ActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing
        pFSel.SelectFeatures pQf, esriSelectionResultNew, True
        m_pMxDoc.ActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing
        
        Call PopulateCommentInfo
        Call PopulateCommodityInfo
        Call PopulateCountyInfo
        Call PopulateMineElevationInfo
        Call PopulateMineInfo
        Call PopulateOperatorInfo
        Call PopulateQuadInfo
        Call PopulateSeamInfo
        Call PopulateTownshipInfo
        
    ElseIf lngCount = 0 Then
        m_pMxDoc.ActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing
        pFSel.Clear
        m_pMxDoc.ActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing
        
        Call ClearCommentInfo
        Call ClearCommodityInfo
        Call ClearCountyInfo
        Call ClearMineElevationInfo
        Call ClearMineInfo
        Call ClearOperatorInfo
        Call ClearQuadInfo
        Call ClearSeamInfo
        Call ClearTownshipInfo
    End If

End Sub
Private Sub PopulateCommentInfo()
'Completed
    Dim pODGSRelClass As ODGSRelClass
    Dim pRelSet As ISet
    Dim pQf As IQueryFilter
    Dim pRow As IRow
    
    Dim lngFldComment As Long
    
    Set pQf = New QueryFilter
    pQf.WhereClause = "MINE_API = '" & m_strMineAPI & "'"
    
    Set m_pLayer = m_pFLayer
    Set pODGSRelClass = New ODGSRelClass
    Set pRelSet = pODGSRelClass.RelClassCursor(m_pLayer, pQf, "COMMENT")
    
    Set pRow = pRelSet.Next
    
    lstComments.Clear
    Do Until pRow Is Nothing
        lngFldComment = pRow.Fields.FindField("CMMNT")
        lstComments.AddItem pRow.value(lngFldComment)
        
        Set pRow = pRelSet.Next
    Loop
End Sub

Private Sub PopulateCommodityInfo()
'Completed
    Dim pODGSRelClass As ODGSRelClass
    Dim pRelSet As ISet
    Dim pQf As IQueryFilter
    Dim pRow As IRow
    
    Dim lngFldCommodity As Long
    
    Set pQf = New QueryFilter
    pQf.WhereClause = "MINE_API = '" & m_strMineAPI & "'"
    
    Set m_pLayer = m_pFLayer
    Set pODGSRelClass = New ODGSRelClass
    Set pRelSet = pODGSRelClass.RelClassCursor(m_pLayer, pQf, "COMMODITY")
    
    Set pRow = pRelSet.Next
    
    lstCommodity.Clear
    Do Until pRow Is Nothing
        lngFldCommodity = pRow.Fields.FindField("COMMODITY")
        lstCommodity.AddItem pRow.value(lngFldCommodity)
        
        Set pRow = pRelSet.Next
    Loop
End Sub

Private Sub PopulateCountyInfo()
'Completed
    Dim pODGSRelClass As ODGSRelClass
    Dim pRelSet As ISet
    Dim pQf As IQueryFilter
    Dim pRow As IRow
    
    Dim lngFldCountyName As Long
    
    Set pQf = New QueryFilter
    pQf.WhereClause = "MINE_API = '" & m_strMineAPI & "'"
    
    Set m_pLayer = m_pFLayer
    Set pODGSRelClass = New ODGSRelClass
    Set pRelSet = pODGSRelClass.RelClassCursor(m_pLayer, pQf, "COUNTY")
    
    Set pRow = pRelSet.Next
    
    lstCounties.Clear
    Do Until pRow Is Nothing
        lngFldCountyName = pRow.Fields.FindField("CTY_NM")
        lstCounties.AddItem pRow.value(lngFldCountyName)
        
        Set pRow = pRelSet.Next
    Loop

End Sub

Private Sub PopulateMineInfo()
    Dim pODGSRelClass As ODGSRelClass
    Dim pRelSet As ISet
    Dim pQf As IQueryFilter
    Dim pFCur As IFeatureCursor
    Dim pRow As IRow
    Dim pRowFC As IRow
    
    Dim lngFldMineType As Long
    Dim lngFldMineNo As Long
    Dim lngFldAbdDate As Long
    Dim lngFldMapDate As Long
    Dim lngFldFromDate As Long
    Dim lngFldToDate As Long
    Dim lngFldLocation As Long
    Dim lngFldOSM As Long
    Dim lngFldOpenningType
    Dim lngFldGISPoly As Long
    Dim lngFldNoPoly As Long
    Dim lngFldActive As Long
    Dim lngFldUTMNorth As Long
    Dim lngFldUTMEast As Long
    Dim lngFldLat As Long
    Dim lngFldLong As Long
    Dim lngFldDrainage As Long
    
    Set pQf = New QueryFilter
    pQf.WhereClause = "MINE_API = '" & m_strMineAPI & "'"
    
    Set m_pLayer = m_pFLayer
    Set pODGSRelClass = New ODGSRelClass
    Set pRelSet = pODGSRelClass.RelClassCursor(m_pLayer, pQf, "MINES")
    
    Set pFCur = m_pFC.Search(pQf, False)
    Set pRowFC = pFCur.NextFeature
    
    txtMineAPI.Text = m_strMineAPI
    
    Set pRow = pRelSet.Next
    
    lngFldMineType = pRow.Fields.FindField("MN_TYPE")
    lngFldMineNo = pRow.Fields.FindField("MN_NO")
    lngFldFromDate = pRow.Fields.FindField("RNG_FRM")
    lngFldToDate = pRow.Fields.FindField("RNG_TO")
    lngFldAbdDate = pRow.Fields.FindField("AB_DT")
    lngFldMapDate = pRow.Fields.FindField("MAP_DT")
    lngFldOSM = pRow.Fields.FindField("OSM_DOC_NO")
    lngFldOpenningType = pRow.Fields.FindField("OPEN_TYPE")
    lngFldLocation = pRow.Fields.FindField("LOCATION")
    lngFldGISPoly = pRow.Fields.FindField("MinePoly")
    lngFldNoPoly = pRow.Fields.FindField("num_polys")
    lngFldActive = pRow.Fields.FindField("Active")
    lngFldUTMNorth = pRow.Fields.FindField("UTM_N")
    lngFldUTMEast = pRow.Fields.FindField("UTM_E")
    lngFldLat = pRow.Fields.FindField("LAT")
    lngFldLong = pRow.Fields.FindField("LONG_")
    
    lngFldDrainage = pRowFC.Fields.FindField("DRAIN")
    
    Do Until pRow Is Nothing
        If VarType(pRow.value(lngFldMineType)) = vbNull Then
            txtMineType.Text = ""
        Else
            txtMineType.Text = pRow.value(lngFldMineType)
        End If
        
        If VarType(pRow.value(lngFldMineNo)) = vbNull Then
            txtMineMapNo.Text = ""
        Else
            txtMineMapNo.Text = pRow.value(lngFldMineNo)
        End If
        
        If VarType(pRow.value(lngFldFromDate)) = vbNull Then
            txtAnnualMapFrom.Text = ""
        Else
            txtAnnualMapFrom.Text = pRow.value(lngFldFromDate)
        End If
        
        If VarType(pRow.value(lngFldToDate)) = vbNull Then
            txtAnnualMapTo.Text = ""
        Else
            txtAnnualMapTo.Text = pRow.value(lngFldToDate)
        End If
            
        If VarType(pRow.value(lngFldAbdDate)) = vbNull Then
            txtAbdDate.Text = ""
        Else
            txtAbdDate.Text = pRow.value(lngFldAbdDate)
        End If
        
        If VarType(pRow.value(lngFldAbdDate)) = vbNull Then
            txtAbdDate2.Text = ""
        Else
            txtAbdDate2.Text = pRow.value(lngFldAbdDate)
        End If
        
        If VarType(pRow.value(lngFldMapDate)) = vbNull Then
            txtMapDate.Text = ""
        Else
            txtMapDate.Text = pRow.value(lngFldMapDate)
        End If
        
        If VarType(pRow.value(lngFldOSM)) = vbNull Then
            txtOSMDoc.Text = ""
        Else
            txtOSMDoc.Text = pRow.value(lngFldOSM)
        End If
            
        If VarType(pRow.value(lngFldOpenningType)) = vbNull Then
            txtOpenning.Text = ""
        Else
            txtOpenning.Text = pRow.value(lngFldOpenningType)
        End If
        
        If VarType(pRow.value(lngFldLocation)) = vbNull Then
            txtLocation.Text = ""
        Else
            txtLocation.Text = pRow.value(lngFldLocation)
        End If
        
        If VarType(pRow.value(lngFldGISPoly)) = vbNull Then
            txtGISPolys.Text = ""
        Else
            txtGISPolys.Text = pRow.value(lngFldGISPoly)
        End If
        
        If VarType(pRow.value(lngFldNoPoly)) = vbNull Then
            txtNumGISPolys.Text = ""
        Else
            txtNumGISPolys.Text = pRow.value(lngFldNoPoly)
        End If
        
        If VarType(pRow.value(lngFldUTMNorth)) = vbNull Then
            txtUTMNorth.Text = ""
        Else
            txtUTMNorth.Text = pRow.value(lngFldUTMNorth)
        End If
        
        If VarType(pRow.value(lngFldUTMEast)) = vbNull Then
            txtUTMEast.Text = ""
        Else
            txtUTMEast.Text = pRow.value(lngFldUTMEast)
        End If

        If VarType(pRow.value(lngFldLat)) = vbNull Then
            txtLatitude.Text = ""
        Else
            txtLatitude.Text = pRow.value(lngFldLat)
        End If
        
        If VarType(pRow.value(lngFldLong)) = vbNull Then
            txtLongitude.Text = ""
        Else
            txtLongitude.Text = pRow.value(lngFldLong)
        End If
        
        If VarType(pRowFC.value(lngFldDrainage)) = vbNull Then
            txtDrainage.Text = ""
        Else
            txtDrainage.Text = pRowFC.value(lngFldDrainage)
        End If

            
        If pRow.value(lngFldActive) = 0 Then
            chkActive.value = False
        Else
            chkActive.value = True
        End If
                
        Set pRow = pRelSet.Next
    Loop

End Sub

Private Sub PopulateMineElevationInfo()
    Dim pODGSRelClass As ODGSRelClass
    Dim pRelSet As ISet
    Dim pQf As IQueryFilter
    Dim pRow As IRow
    
    Dim lngFldElevation As Long
    
    Set pQf = New QueryFilter
    pQf.WhereClause = "MINE_API = '" & m_strMineAPI & "'"
    
    Set m_pLayer = m_pFLayer
    Set pODGSRelClass = New ODGSRelClass
    Set pRelSet = pODGSRelClass.RelClassCursor(m_pLayer, pQf, "ELEVATION")
    
    Set pRow = pRelSet.Next
    
    Do Until pRow Is Nothing
        lngFldElevation = pRow.Fields.FindField("ELEV")
        If VarType(pRow.value(lngFldElevation)) = vbNull Then
            txtElevation.Text = ""
        Else
            txtElevation.Text = pRow.value(lngFldElevation)
        End If
        
        Set pRow = pRelSet.Next
    Loop

End Sub

Private Sub PopulateOperatorInfo()
'Completed
    Dim pODGSRelClass As ODGSRelClass
    Dim pRelSet As ISet
    Dim pQf As IQueryFilter
    Dim pRow As IRow
    
    Dim i As Long
    Dim lngFldOperator As Long
    Dim lngFldMineName As Long
    Dim lngFldMostRecentName As Long
        
    Dim strOperator As String
    Dim strMineName As String
    Dim lngMostRecentName As Long
    
    Set pQf = New QueryFilter
    pQf.WhereClause = "MINE_API = '" & m_strMineAPI & "'"
    
    Set m_pLayer = m_pFLayer
    Set pODGSRelClass = New ODGSRelClass
    Set pRelSet = pODGSRelClass.RelClassCursor(m_pLayer, pQf, "OPERATOR")
    
    Set pRow = pRelSet.Next
    
    lstOperatorMineName.Clear
    
    Do Until pRow Is Nothing
        lngFldOperator = pRow.Fields.FindField("OP_NAME")
        lngFldMineName = pRow.Fields.FindField("MN_NAME")
        lngFldMostRecentName = pRow.Fields.FindField("PRESENT")
        
        'Get the operator name
        If VarType(pRow.value(lngFldOperator)) = vbNull Then
            strOperator = ""
        Else
            strOperator = pRow.value(lngFldOperator)
        End If
        
        'Get the mine name
        If VarType(pRow.value(lngFldMineName)) = vbNull Then
            strMineName = ""
        Else
            strMineName = pRow.value(lngFldMineName)
        End If
        
        'Determine if the name of the mine and operator were at abandonment
        If VarType(pRow.value(lngFldMostRecentName)) = vbNull Then
            lngMostRecentName = 0
        ElseIf pRow.value(lngFldMostRecentName) = 1 Then
            lngMostRecentName = 1
        Else
            lngMostRecentName = 0
        End If
        
        'Populate the row in the list box
        If lngMostRecentName = 0 Then
            If strMineName = "" Then
                lstOperatorMineName.AddItem strOperator
            Else
                lstOperatorMineName.AddItem strOperator & "/" & strMineName
            End If
        ElseIf lngMostRecentName = 1 Then
            If strMineName = "" Then
                lstOperatorMineName.AddItem strOperator & "/*"
            Else
                lstOperatorMineName.AddItem strOperator & "/" & strMineName & "/*"
            End If
        End If

        Set pRow = pRelSet.Next
    Loop
    
'    lstOperatorMineName.List(0, 0) = "OPERATOR"
'    lstOperatorMineName.List(1, 0) = "MINE NAME"
'    i = 1
    
'    Do Until pRow Is Nothing
'        If VarType(pRow.value(lngFldOperator)) = vbNull Then
'            lstOperatorMineName.List(i, 0) = ""
'        Else
'            lstOperatorMineName.List(i, 0) = pRow.value(lngFldOperator)
'        End If
'
'        If VarType(pRow.value(lngFldMineName)) = vbNull Then
'            lstOperatorMineName.List(i, 1) = ""
'        Else
'            lstOperatorMineName.List(i, 1) = pRow.value(lngFldMineName)
'        End If
'
'        Set pRow = pRelSet.Next
'        i = i + 1
'    Loop
End Sub

Private Sub PopulateQuadInfo()
'Completed
    Dim pODGSRelClass As ODGSRelClass
    Dim pRelSet As ISet
    Dim pQf As IQueryFilter
    Dim pRow As IRow
    
    Dim lngFldQuad24K As Long
    
    Set pQf = New QueryFilter
    pQf.WhereClause = "MINE_API = '" & m_strMineAPI & "'"
    
    Set m_pLayer = m_pFLayer
    Set pODGSRelClass = New ODGSRelClass
    Set pRelSet = pODGSRelClass.RelClassCursor(m_pLayer, pQf, "QUAD")
    
    Set pRow = pRelSet.Next
    
    lstQuad24K.Clear
    Do Until pRow Is Nothing
        lngFldQuad24K = pRow.Fields.FindField("QUAD_NM")
        lstQuad24K.AddItem pRow.value(lngFldQuad24K)
        
        Set pRow = pRelSet.Next
    Loop

End Sub

Private Sub PopulateSeamInfo()
'Completed
    Dim pODGSRelClass As ODGSRelClass
    Dim pRelSet As ISet
    Dim pQf As IQueryFilter
    Dim pRow As IRow
    
    Dim lngFldSeam As Long
    
    Set pQf = New QueryFilter
    pQf.WhereClause = "MINE_API = '" & m_strMineAPI & "'"
    
    Set m_pLayer = m_pFLayer
    Set pODGSRelClass = New ODGSRelClass
    Set pRelSet = pODGSRelClass.RelClassCursor(m_pLayer, pQf, "SEAM")
    
    Set pRow = pRelSet.Next
    
    lstSeam.Clear
    Do Until pRow Is Nothing
        lngFldSeam = pRow.Fields.FindField("Coal Bed")
        If VarType(pRow.value(lngFldSeam)) = vbNull Then
            lstSeam.AddItem ""
        Else
            lstSeam.AddItem pRow.value(lngFldSeam)
        End If
        
        Set pRow = pRelSet.Next
    Loop
    
End Sub

Private Sub PopulateTownshipInfo()
'Completed
    Dim pODGSRelClass As ODGSRelClass
    Dim pRelSet As ISet
    Dim pQf As IQueryFilter
    Dim pRow As IRow
    
    Dim lngFldTownship As Long
    
    Set pQf = New QueryFilter
    pQf.WhereClause = "MINE_API = '" & m_strMineAPI & "'"
    
    Set m_pLayer = m_pFLayer
    Set pODGSRelClass = New ODGSRelClass
    Set pRelSet = pODGSRelClass.RelClassCursor(m_pLayer, pQf, "TOWNSHIP")
    
    Set pRow = pRelSet.Next
    
    lstTownships.Clear
    Do Until pRow Is Nothing
        lngFldTownship = pRow.Fields.FindField("TWP_NAME")
        lstTownships.AddItem pRow.value(lngFldTownship)
        
        Set pRow = pRelSet.Next
    Loop
End Sub

Private Sub ClearCommentInfo()
'Completed
    lstComments.Clear
End Sub

Private Sub ClearCommodityInfo()
'Completed
    lstCommodity.Clear
End Sub

Private Sub ClearCountyInfo()
'Completed
    lstCounties.Clear

End Sub

Private Sub ClearMineInfo()
    txtMineAPI = ""
    txtMineMapNo = ""
    txtMineType.Text = ""
    txtAnnualMapFrom.Text = ""
    txtAnnualMapTo.Text = ""
    txtAbdDate.Text = ""
    txtAbdDate2.Text = ""
    txtMapDate.Text = ""
    txtOSMDoc.Text = ""
    txtOpenning.Text = ""
    txtLocation.Text = ""
    txtGISPolys.Text = ""
    txtNumGISPolys.Text = ""
    txtUTMNorth.Text = ""
    txtUTMEast.Text = ""
    txtLatitude.Text = ""
    txtLongitude.Text = ""
    txtDrainage.Text = ""
    chkActive.value = False
    
End Sub

Private Sub ClearMineElevationInfo()
'Completed
    txtElevation.Text = ""

End Sub

Private Sub ClearOperatorInfo()
'Completed
    lstOperatorMineName.Clear

End Sub

Private Sub ClearQuadInfo()
'Completed
    lstQuad24K.Clear

End Sub

Private Sub ClearSeamInfo()
'Completed
    lstSeam.Clear

End Sub

Private Sub ClearTownshipInfo()
'Completed
    lstTownships.Clear
    
End Sub

Private Sub cmdQuit_Click()
'Completed
    Dim pFSel As IFeatureSelection
    
    m_pMxDoc.ActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing
    Set pFSel = m_pFLayer
    pFSel.Clear
    m_pMxDoc.ActiveView.PartialRefresh esriViewGeoSelection, Nothing, Nothing

    Call ExecuteCmd
    Unload Me
    
End Sub

Private Sub cmdMineMapImages_Click()
    Dim pNewLayer As ILayer
    Dim pGrpLayer As IGroupLayer
    
    Dim i As Long
    Dim lngFldMineAPI As Long
    Dim lngFldPath As Long
    Dim lngFldFileName As Long
    
    Dim pStdTableColl As IStandaloneTableCollection
    Dim pStdTable As IStandaloneTable
    Dim pTable As ITable
    
    Dim pQf As IQueryFilter
    Dim pCursor As ICursor
    Dim pRow As IRow
    
    Dim pODGSMapScanLyr As ODGSMapScanLayer
    
    Dim strFolder As String
    Dim strFileName As String
        
    'START
    'Find the AUM Mine Map Images metadata table
    Set pStdTableColl = m_pMap
    For i = 0 To pStdTableColl.StandaloneTableCount - 1
        Set pStdTable = pStdTableColl.StandaloneTable(i)
        If pStdTable.Name = "AUMMINEMAPS" Then
            Set pTable = pStdTable
            lngFldMineAPI = pTable.FindField("MINE_API")
            lngFldPath = pTable.FindField("PATH")
            lngFldFileName = pTable.FindField("FILENAME")
        End If
    Next i
    
    If pTable Is Nothing Then Exit Sub 'If the metadata table has been detached or deleted from the ArcMap, exit the commanad.
    
    'Create new group layer
    Set pGrpLayer = New GroupLayer
    pGrpLayer.Name = "AUM Georeferenced Mine Maps - " & m_strMineAPI
    pGrpLayer.Expanded = True
    pGrpLayer.Visible = False
    
    'Create a cursor on the AUM Mine Map Images metadata table
    Set pQf = New QueryFilter
    pQf.WhereClause = "[" & pTable.Fields.Field(lngFldMineAPI).Name & "] = '" & m_strMineAPI & "'"
    
    If pTable.RowCount(pQf) <= 0 Then Exit Sub 'If there are no georeferenced mine map images, exit the command.
    
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

End Sub

Private Sub LoadGeoRefMineMaps(pGeom As IGeometry)
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

