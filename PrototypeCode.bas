Attribute VB_Name = "PrototypeCode"
Option Explicit

Public Sub TestAddLayers()
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pFlayer As IFeatureLayer
    Dim pLyrPos As ILayerPosition

    Dim pGxLayer As IGxLayer, pGxNewLayer As IGxLayer, pGxFile As IGxFile
    Dim pGrpLayer As IGroupLayer
    
    'Start
    Set pMxDoc = ThisDocument
    Set pMap = pMxDoc.FocusMap
    
    Set pGrpLayer = New GroupLayer
'    pGrpLayer.Name = "AUM Layers"
    
    Set pGxLayer = New esriCatalog.GxLayer
    Set pGxFile = pGxLayer
    pGxFile.Path = "D:\OMSIUA\Layer_Files\Mine Opening from Topographic Maps.lyr"
    pGxFile.Open
    Set pLayer = pGxLayer.Layer
        Set pFlayer = pLayer
        Set pLyrPos = pFlayer
        pLyrPos.LayerWeight = 10
    pMap.AddLayer pLayer
'    pGrpLayer.Add pLayer
    
    Set pGxLayer = New esriCatalog.GxLayer
    Set pGxFile = pGxLayer
    pGxFile.Path = "D:\OMSIUA\Layer_Files\Mine Opening.lyr"
    pGxFile.Open
    Set pLayer = pGxLayer.Layer
            Set pFlayer = pLayer
        Set pLyrPos = pFlayer
        pLyrPos.LayerWeight = 9
    pMap.AddLayer pLayer

'    pGrpLayer.Add pLayer
    
    Set pGxLayer = New esriCatalog.GxLayer
    Set pGxFile = pGxLayer
    pGxFile.Path = "D:\OMSIUA\Layer_Files\Mine Location - Extent Unknown.lyr"
    pGxFile.Open
    Set pLayer = pGxLayer.Layer
            Set pFlayer = pLayer
        Set pLyrPos = pFlayer
        pLyrPos.LayerWeight = 8
    pMap.AddLayer pLayer

'    pGrpLayer.Add pLayer
    
    Set pGxLayer = New esriCatalog.GxLayer
    Set pGxFile = pGxLayer
    pGxFile.Path = "D:\OMSIUA\Layer_Files\Underground Mine - Extent Partially Unknown.lyr"
    pGxFile.Open
    Set pLayer = pGxLayer.Layer
            Set pFlayer = pLayer
        Set pLyrPos = pFlayer
        pLyrPos.LayerWeight = 7
    pMap.AddLayer pLayer

'    pGrpLayer.Add pLayer
    
    Set pGxLayer = New esriCatalog.GxLayer
    Set pGxFile = pGxLayer
    pGxFile.Path = "D:\OMSIUA\Layer_Files\Superimposed Underground Mine.lyr"
    pGxFile.Open
    Set pLayer = pGxLayer.Layer
            Set pFlayer = pLayer
        Set pLyrPos = pFlayer
        pLyrPos.LayerWeight = 6
    pMap.AddLayer pLayer

'    pGrpLayer.Add pLayer
    
    Set pGxLayer = New esriCatalog.GxLayer
    Set pGxFile = pGxLayer
    pGxFile.Path = "D:\OMSIUA\Layer_Files\Underground Mine.lyr"
    pGxFile.Open
    Set pLayer = pGxLayer.Layer
        Set pFlayer = pLayer
        Set pLyrPos = pFlayer
        pLyrPos.LayerWeight = 5
    pMap.AddLayer pLayer
'    pGrpLayer.Add pLayer
    
    
    pMxDoc.ActiveView.Refresh
    
'    Set pEnumLayer = pMap.Layers
'    Set pLayer = pEnumLayer.Next
'    Do Until pLayer Is Nothing
'        If pLayer.Name = "Underground Mine" Then
'            Set pFLayer = pLayer
'            Set pLyrPos = pFLayer
'            pLyrPos.LayerWeight = 1
'        End If
'        Set pLayer = pEnumLayer.Next
'    Loop
    
'    pMxDoc.ActiveView.Refresh
    
End Sub

Public Sub MapScan()

    Dim pGxObj As IGxObject
    Dim pGxFile As IGxFile
    Dim blnExist As Boolean
    strDocFolder As String
    strDocFileName As String
    
    Set pGxFile = New GxFile
    pGxFile.Path = strDocFolder & "\" & strDocFileName
    Set pGxObj = pGxFile
    
    If pGxObj.IsValid = False Then
        blnExist = False
    Else
        blnExist = True
    End If
    
End Sub

Public Property Get Exist() As Boolean
On Error GoTo Exist_Err

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

Public Sub QueryTableJoins()
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pFlayer As IFeatureLayer
    Dim pGFLyr As IGeoFeatureLayer
    Dim pDS As IDataset
    Dim pFC As IFeatureClass
    Dim pTable As ITable
    Dim pFields As IFields
    Dim pField As IField
    
    Dim pQf As IQueryFilter
    Dim pFCur As IFeatureCursor
    Dim pF As IFeature
    
    Dim i As Long
    Dim lngFldCnt As Long
    Dim lngFldFullName As Long
        
    'START
    Set pMxDoc = ThisDocument
    Set pMap = pMxDoc.FocusMap
    Set pEnumLayer = pMap.Layers
    Set pLayer = pEnumLayer.Next
    
    Do Until pLayer Is Nothing
        If pLayer.Name = "NCRDS_PTS" Then
            Set pFlayer = pLayer
            Set pGFLyr = pFlayer
'            Set pFC = pFLayer.FeatureClass
            Set pFC = pGFLyr.DisplayFeatureClass
            Set pDS = pFC
            Set pTable = pFC
            Set pFields = pTable.Fields
            lngFldCnt = pFields.FieldCount
        End If
        Set pLayer = pEnumLayer.Next

    Loop
    
    Set pLayer = pGFLyr
    Debug.Print "Layer Name = " & pLayer.Name
    Debug.Print vbTab & "Feature Class Name = " & pDS.Name
    
    For i = 0 To lngFldCnt - 1
        Set pField = pFC.Fields.Field(i)
        Debug.Print vbTab; pField.Name
    Next i
    
    lngFldFullName = pFC.FindField("ODGSDOCLOCATIONS.FULLNAME")
    
    'Perform Query
    Set pQf = New QueryFilter
    pQf.WhereClause = "ODGSDOCLOCATIONS.ODGSDOCID LIKE " & "'C013*'"
    
    Debug.Print pFC.FeatureCount(pQf)
    
    Set pFCur = pGFLyr.SearchDisplayFeatures(pQf, True)
    Set pF = pFCur.NextFeature
    
    Do Until pF Is Nothing
        Debug.Print vbTab & vbTab & pF.value(lngFldFullName)
        Set pF = pFCur.NextFeature
    Loop
    
End Sub

Public Sub FindJoins()
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pFlayer As IFeatureLayer
    Dim pGFLyr As IGeoFeatureLayer
    Dim pDRelClass As IDisplayRelationshipClass
    Dim pRelClass As IRelationshipClass
    Dim pDS As IDataset
    Dim pDSDest As IDataset
    Dim pDSOrigin As IDataset
    Dim pFC As IFeatureClass
    Dim pTable As ITable
    Dim pFields As IFields
    Dim pField As IField
    
    Dim pQf As IQueryFilter
    Dim pFCur As IFeatureCursor
    Dim pF As IFeature
    
    Dim i As Long
    Dim lngFldCnt As Long
    Dim lngFldFullName As Long
    
    'START
    Set pMxDoc = ThisDocument
    Set pMap = pMxDoc.FocusMap
    Set pEnumLayer = pMap.Layers
    Set pLayer = pEnumLayer.Next
    
    Do Until pLayer Is Nothing
        If pLayer.Name = "Underground Mine" Then
            Set pFlayer = pLayer
            Set pDRelClass = pFlayer
            Set pRelClass = pDRelClass.RelationshipClass
            Set pDS = pRelClass
            Set pDSDest = pRelClass.DestinationClass
            Set pDSOrigin = pRelClass.OriginClass
'            Set pGFLyr = pFLayer
'            Set pFC = pFLayer.FeatureClass
'            Set pFC = pGFLyr.DisplayFeatureClass
'            Set pDS = pFC
'            Set pTable = pFC
'            Set pFields = pTable.Fields
'            lngFldCnt = pFields.FieldCount
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    
    
    Debug.Print pDS.BrowseName
    Debug.Print pRelClass.BackwardPathLabel
    Debug.Print pRelClass.Cardinality
    Debug.Print pRelClass.DestinationForeignKey
    Debug.Print pRelClass.DestinationPrimaryKey
    Debug.Print pRelClass.ForwardPathLabel
    Debug.Print pRelClass.IsAttributed
    Debug.Print pRelClass.IsComposite
    Debug.Print pRelClass.Notification
    Debug.Print pRelClass.OriginForeignKey
    Debug.Print pRelClass.OriginPrimaryKey
    Debug.Print pRelClass.RelationshipClassID
'    Debug.Print pRelClass.RelationshipRules

End Sub

Public Sub AddJoin()
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pFlayer As IFeatureLayer
    Dim pGFLayer As IGeoFeatureLayer
    Dim pFC As IFeatureClass
    
    Dim pStdTblColl As IStandaloneTableCollection
    Dim pStdTable As IStandaloneTable
    Dim pTable As ITable
    Dim pDTable As IDisplayTable
    
    Dim i As Long
    Dim lngTableCount As Long
    
    Dim strJnFld As String
    
    Dim pMemRelFact As IMemoryRelationshipClassFactory
    Dim pRelClass As IRelationshipClass
    Dim pDisRelClass As IDisplayRelationshipClass
            
    'START
    Set pMxDoc = ThisDocument
    Set pMap = pMxDoc.FocusMap
    Set pEnumLayer = pMap.Layers
    Set pLayer = pEnumLayer.Next
    
    'Find the Feature Class
    Do Until pLayer Is Nothing
        If pLayer.Name = "Underground Mine" Then
            Set pFlayer = pLayer
            Set pGFLayer = pFlayer
            Set pFC = pGFLayer.DisplayFeatureClass
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    
    'Find the Standalong Table that will be joined to the Feature Class
    Set pStdTblColl = pMap
    lngTableCount = pStdTblColl.StandaloneTableCount
    
    For i = 0 To lngTableCount - 1
        Set pStdTable = pStdTblColl.StandaloneTable(i)
        If pStdTable.Name = "GEOLOGY.LOADER.AUM_TBLMINES" Then
            Set pDTable = pStdTable
            Set pTable = pDTable.DisplayTable
        End If
    Next i
    
    'Join the table
    strJnFld = "MINE_API"
    
    Set pMemRelFact = New MemoryRelationshipClassFactory
    
    Set pRelClass = pMemRelFact.Open("TestJoin", pFC, strJnFld, pTable, strJnFld, "NCRDS", "Document", esriRelCardinalityOneToOne)
    
    Debug.Print "pRelClass"
    Debug.Print vbTab & pRelClass.BackwardPathLabel
    Debug.Print vbTab & pRelClass.Cardinality
    Debug.Print vbTab & pRelClass.DestinationForeignKey
    Debug.Print vbTab & pRelClass.DestinationPrimaryKey
    Debug.Print vbTab & pRelClass.ForwardPathLabel
    Debug.Print vbTab & pRelClass.IsAttributed
    Debug.Print vbTab & pRelClass.IsComposite
    Debug.Print vbTab & pRelClass.Notification
    Debug.Print vbTab & pRelClass.OriginForeignKey
    Debug.Print vbTab & pRelClass.OriginPrimaryKey
    Debug.Print vbTab & pRelClass.RelationshipClassID
    
    Set pDisRelClass = pFlayer
    pDisRelClass.DisplayRelationshipClass pRelClass, esriLeftOuterJoin

End Sub

Public Sub AddMultipleJoins()
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pFlayer As IFeatureLayer
    Dim pGFLayer As IGeoFeatureLayer
    Dim pDS1 As IDataset
    Dim pDS2 As IDataset
    Dim pDS3 As IDataset
    Dim pFC As IFeatureClass
    
    Dim pStdTblColl As IStandaloneTableCollection
    Dim pStdTable1 As IStandaloneTable
    Dim pStdTable2 As IStandaloneTable
    Dim pTable1 As ITable
    Dim pTable2 As ITable
    Dim pDTable1 As IDisplayTable
    Dim pDTable2 As IDisplayTable
    
    Dim i As Long
    Dim lngTableCount As Long
    
    Dim strJnFld As String
    
    Dim pMemRelFact1 As IMemoryRelationshipClassFactory
    Dim pMemRelFact2 As IMemoryRelationshipClassFactory
    Dim pRelClass1 As IRelationshipClass
    Dim pRelClass2 As IRelationshipClass
    Dim pDisRelClass As IDisplayRelationshipClass
    
    Dim pRelQryTblFact As IRelQueryTableFactory
    Dim pRelQryTable1 As IRelQueryTable
    Dim pRelQryTable2 As IRelQueryTable
    
    Dim pCur As ICursor
    Dim pFields As IFields
    Dim pField As IField
    Dim pRow As IRow
    
    'START
    Set pMxDoc = ThisDocument
    Set pMap = pMxDoc.FocusMap
    Set pEnumLayer = pMap.Layers
    Set pLayer = pEnumLayer.Next
    
    'Find the Feature Class
    Do Until pLayer Is Nothing
        If pLayer.Name = "Underground Mine" Then
            Set pFlayer = pLayer
            Set pGFLayer = pFlayer
            Set pFC = pGFLayer.DisplayFeatureClass
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    
    'Find the First Standalong Table that will be joined to the Feature Class
    Set pStdTblColl = pMap
    lngTableCount = pStdTblColl.StandaloneTableCount
    
    For i = 0 To lngTableCount - 1
        Set pStdTable1 = pStdTblColl.StandaloneTable(i)
        If pStdTable1.Name = "TBLMINES" Then
            Set pDTable1 = pStdTable1
            Set pTable1 = pDTable1.DisplayTable
        End If
    Next i
    
    'Join the table
    strJnFld = "MINE_API"
    
    Set pMemRelFact1 = New MemoryRelationshipClassFactory
    
    Set pRelClass1 = pMemRelFact1.Open("TestJoin", pFC, strJnFld, pTable1, strJnFld, "Original FC", "First Table", esriRelCardinalityOneToOne)
    Set pDS1 = pRelClass1
    
'    Debug.Print "pRelClass1"
'    Debug.Print pDS1.Name
'    Debug.Print vbTab & pRelClass1.BackwardPathLabel
'    Debug.Print vbTab & pRelClass1.Cardinality
'    Debug.Print vbTab & pRelClass1.DestinationForeignKey
'    Debug.Print vbTab & pRelClass1.DestinationPrimaryKey
'    Debug.Print vbTab & pRelClass1.ForwardPathLabel
'    Debug.Print vbTab & pRelClass1.IsAttributed
'    Debug.Print vbTab & pRelClass1.IsComposite
'    Debug.Print vbTab & pRelClass1.Notification
'    Debug.Print vbTab & pRelClass1.OriginForeignKey
'    Debug.Print vbTab & pRelClass1.OriginPrimaryKey
'    Debug.Print vbTab & pRelClass1.RelationshipClassID
'    Debug.Print
    
    'Create the first relquerytable
    Set pRelQryTblFact = New RelQueryTableFactory
    Set pRelQryTable1 = pRelQryTblFact.Open(pRelClass1, True, Nothing, Nothing, "", True, True)
    
'    Set pCur = pRelQryTable1.DestinationTable.Search(Nothing, False)
'    Set pFields = pRelQryTable1.DestinationTable.Fields
'
'    Debug.Print "pRelQryTable1.DestinationTable Field Names"
'    For i = 0 To pFields.FieldCount - 1
'        Debug.Print vbTab & pFields.Field(i).Name
'    Next i
'    Debug.Print
'
'    Set pCur = pRelQryTable1.SourceTable.Search(Nothing, False)
'    Set pFields = pRelQryTable1.SourceTable.Fields
'
'    Debug.Print "pRelQryTable1.SourceTable Field Names"
'    For i = 0 To pFields.FieldCount - 1
'        Debug.Print vbTab & pFields.Field(i).Name
'    Next i
'    Debug.Print
'
'    Debug.Print "pRelQryTable1.RelationshipClass Attributes"
'    Debug.Print vbTab & pRelQryTable1.RelationshipClass.BackwardPathLabel
'    Debug.Print vbTab & pRelQryTable1.RelationshipClass.Cardinality
'    Debug.Print vbTab & pRelQryTable1.RelationshipClass.DestinationForeignKey
'    Debug.Print vbTab & pRelQryTable1.RelationshipClass.DestinationPrimaryKey
'    Debug.Print vbTab & pRelQryTable1.RelationshipClass.ForwardPathLabel
'    Debug.Print vbTab & pRelQryTable1.RelationshipClass.IsAttributed
'    Debug.Print vbTab & pRelQryTable1.RelationshipClass.IsComposite
'    Debug.Print vbTab & pRelQryTable1.RelationshipClass.Notification
'    Debug.Print vbTab & pRelQryTable1.RelationshipClass.OriginForeignKey
'    Debug.Print vbTab & pRelQryTable1.RelationshipClass.OriginPrimaryKey
'    Debug.Print vbTab & pRelQryTable1.RelationshipClass.RelationshipClassID
'    Debug.Print
    
    'Find the Second Standalong Table that will be joined to the Feature Class
'    Set pStdTblColl = pMap
'    lngTableCount = pStdTblColl.StandaloneTableCount
    
    For i = 0 To lngTableCount - 1
        Set pStdTable2 = pStdTblColl.StandaloneTable(i)
        If pStdTable2.Name = "TBLOPERATOR" Then
            Set pDTable2 = pStdTable2
            Set pTable2 = pDTable2.DisplayTable
        End If
    Next i
    
    Set pCur = pTable2.Search(Nothing, False)
    Set pFields = pTable2.Fields
    
'    Debug.Print "2nd Destination Table - From Display Table"
'    For i = 0 To pFields.FieldCount - 1
'        Debug.Print vbTab & pFields.Field(i).Name
'    Next i
    
    'Join the table
    strJnFld = "MINE_API"
    
    Set pMemRelFact2 = New MemoryRelationshipClassFactory
    
    Set pRelClass2 = pMemRelFact2.Open("TestJoin2", pRelQryTable1, "geology.LOADER.UNDERGROUND_MINES.MINE_API", pTable2, strJnFld, "OrigRelClass", "SecondTable", esriRelCardinalityOneToOne)
    
'    Debug.Print "pRelClass2"
'    Debug.Print vbTab & pRelClass2.BackwardPathLabel
'    Debug.Print vbTab & pRelClass2.Cardinality
'    Debug.Print vbTab & pRelClass2.DestinationForeignKey
'    Debug.Print vbTab & pRelClass2.DestinationPrimaryKey
'    Debug.Print vbTab & pRelClass2.ForwardPathLabel
'    Debug.Print vbTab & pRelClass2.IsAttributed
'    Debug.Print vbTab & pRelClass2.IsComposite
'    Debug.Print vbTab & pRelClass2.Notification
'    Debug.Print vbTab & pRelClass2.OriginForeignKey
'    Debug.Print vbTab & pRelClass2.OriginPrimaryKey
'    Debug.Print vbTab & pRelClass2.RelationshipClassID
    
    'Create the second relquerytable
    Set pRelQryTblFact = New RelQueryTableFactory
    Set pRelQryTable2 = pRelQryTblFact.Open(pRelClass2, True, Nothing, Nothing, "", True, True)

    'Display the Join in the Feature Class layer
    Set pDisRelClass = pFlayer
    pDisRelClass.DisplayRelationshipClass pRelQryTable2.RelationshipClass, esriLeftOuterJoin
End Sub

Public Sub AddRelate()
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pFlayer As IFeatureLayer
    Dim pGFLayer As IGeoFeatureLayer
    Dim pFC As IFeatureClass
    
    Dim pStdTblColl As IStandaloneTableCollection
    Dim pStdTable As IStandaloneTable
    Dim pTable As ITable
    Dim pDTable As IDisplayTable
    
    Dim i As Long
    Dim lngTableCount As Long
    
    Dim strJnFld As String
    
    Dim pMemRelFact As IMemoryRelationshipClassFactory
    Dim pRelClass As IRelationshipClass
    Dim pRelClassColl As IRelationshipClassCollectionEdit
    
    Dim pDisRelClass As IDisplayRelationshipClass
            
    'START
    Set pMxDoc = ThisDocument
    Set pMap = pMxDoc.FocusMap
    Set pEnumLayer = pMap.Layers
    Set pLayer = pEnumLayer.Next
    
    'Find the Feature Class
    Do Until pLayer Is Nothing
        If pLayer.Name = "Underground Mine" Then
            Set pFlayer = pLayer
            Set pGFLayer = pFlayer
            Set pFC = pGFLayer.DisplayFeatureClass
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    
    'Find the Standalong Table that will be joined to the Feature Class
    Set pStdTblColl = pMap
    lngTableCount = pStdTblColl.StandaloneTableCount
    
    For i = 0 To lngTableCount - 1
        Set pStdTable = pStdTblColl.StandaloneTable(i)
        If pStdTable.Name = "TBLOPERATOR" Then
            Set pDTable = pStdTable
            Set pTable = pDTable.DisplayTable
        End If
    Next i
    
    'Join the table
    strJnFld = "MINE_API"
    
    Set pMemRelFact = New MemoryRelationshipClassFactory
    Set pRelClass = pMemRelFact.Open("TestRelate", pFC, strJnFld, pTable, strJnFld, "Underground Mines", "MineInfo Table", esriRelCardinalityOneToMany)
    
    Set pRelClassColl = pFlayer
    pRelClassColl.AddRelationshipClass pRelClass
End Sub

Public Sub QueryRelate()
    'This will label every selected feature in the polygon coverage
    'or all features if no features are selected
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pEnumLayer As IEnumLayer
    Dim pLayer As ILayer
    Dim pFlayer As IFeatureLayer
    Dim pFSel As IFeatureSelection
    Dim pSelSet As ISelectionSet
    Dim pGFLayer As IGeoFeatureLayer
    Dim pRelClassColl As IRelationshipClassCollection
    Dim pFC As IFeatureClass
    
    Dim pQf As IQueryFilter
    Dim pFCur As IFeatureCursor
    Dim pF As IFeature
    
    Dim pEnumRelClass As IEnumRelationshipClass
    Dim pRelClass As IRelationshipClass
    Dim pDS As IDataset
    
    Dim lngFCount As Long

    Dim i As Long
    Dim r As Long
    Dim myCount As Integer
    Dim pRelObjSet As ISet
    Dim lngFieldIndex As Long
    Dim lngFieldIndex2 As Long
    Dim pObjectClass As IObjectClass
    
    'START
    Set pMxDoc = ThisDocument
    Set pMap = pMxDoc.FocusMap
    Set pEnumLayer = pMap.Layers
    
    Set pLayer = pEnumLayer.Next
    Do Until pLayer Is Nothing
        If pLayer.Name = "Underground Mine" Then
            Set pFlayer = pLayer
            Set pGFLayer = pFlayer
            Set pRelClassColl = pFlayer
'            Set pFC = pGFLayer.DisplayFeatureClass
        End If
        Set pLayer = pEnumLayer.Next
    Loop
    
    
    'Get Selected Features from the Feature Layer
    Set pFSel = pFlayer
    Set pSelSet = pFSel.SelectionSet
    
    Set pQf = New QueryFilter
    pQf.WhereClause = ""
    If pSelSet.Count < 1 Then
        MsgBox "Please select one or more underground mines", vbCritical, "No Selected Mines"
        Exit Sub
    Else
        'Retrieve just the selected features
        'This statement moves the selected features into a feature cursor.  All
        'subsequent work is done using the feature cursor.
        pSelSet.Search pQf, False, pFCur
    End If
    Set pF = pFCur.NextFeature
    
    'Get the Relationship Classes
    Set pEnumRelClass = pRelClassColl.RelationshipClasses
    Set pRelClass = pEnumRelClass.Next
  
    If pRelClass Is Nothing Then
        MsgBox "     No Relationship Class created for this layer." & _
        vbNewLine & "Please create a Relationship Class in ArcCatalog", _
        vbOKOnly, "Relationship Class Missing"
        Exit Sub
    End If
    
    Do Until pRelClass Is Nothing
        Set pDS = pRelClass
        If pDS.BrowseName = "OPERATOR" Then
            Do Until pF Is Nothing
            Debug.Print pDS.BrowseName

                'Now loop through using GetObjectsRelatedToObject

                Dim pRelRow As IRow
                Set pRelObjSet = pRelClass.GetObjectsRelatedToObject(pF)
            
                'Get the destination field
                Set pObjectClass = pRelClass.DestinationClass
                lngFieldIndex = pObjectClass.FindField("OP_NAME")
                lngFieldIndex2 = pObjectClass.FindField("MN_NAME")
                
                Set pRelRow = pRelObjSet.Next
                Do Until pRelRow Is Nothing
                    Debug.Print vbTab & "Operator = " & pRelRow.value(lngFieldIndex)
                    Debug.Print vbTab & "Mine Name = " & pRelRow.value(lngFieldIndex2)
                    Set pRelRow = pRelObjSet.Next
                Loop
                
                Set pF = pFCur.NextFeature
            Loop
        End If
        
        Set pRelClass = pEnumRelClass.Next
    Loop
   

End Sub

Public Sub StackMapStickFigures()
    'This will label every selected feature in the polygon coverage
    'or all features if no features are selected
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
    Dim pLayer As ILayer
    Dim pFlayer As IFeatureLayer
    Dim pFClass As IFeatureClass
    Dim pFieldName As String
    Dim pFieldName2 As String
    Dim pQFilt As IQueryFilter
    Dim pFeature As IFeature
    Dim pFeatCursor As IFeatureCursor
    Dim pFeatureSelection As IFeatureSelection
    Dim pSelectionSet As ISelectionSet
    Dim pIndex As Long
    Dim pFeatCount As Long
      
    Set pMxDoc = ThisDocument
    Set pMap = pMxDoc.FocusMap
    
    pFieldName = "INDICT" 'Enter field name here
    pFieldName2 = "GEO_DEPTH_TO"
  
    'Must have the feature class added to ArcMap as the first layer
    If pMap.LayerCount < 1 Then
        MsgBox "No layer in TOC. Please add a polygon layer.", _
        vbOKOnly, "Missing layer"
        Exit Sub
    End If
  
    Set pLayer = pMap.Layer(0)
    Set pFlayer = pLayer
    
    'Open the FeatureClass
    Set pFClass = pFlayer.FeatureClass
    
    'Labels all features with each features related rows
    'Get any selected features
    Set pFeatureSelection = pFlayer 'QI
    Set pSelectionSet = pFeatureSelection.SelectionSet
    'Create an empty query filter
    Set pQFilt = New QueryFilter
    pQFilt.WhereClause = ""
    If pSelectionSet.Count < 1 Then
        MsgBox "Please select one or more wells", vbCritical, "No Selected Wells"
        Exit Sub
        'Retrieve all features
        'Set pFeatCursor = pFClass.Search(pQFilt, False)
    Else
        'Retrieve just the selected features
        'This statement moves the selected features into a feature cursor.  All
        'subsequent work is done using the feature cursor.
        pSelectionSet.Search pQFilt, False, pFeatCursor
    End If
    
    'Get the first feature
    Set pFeature = pFeatCursor.NextFeature
    
    'Get the first Relationship Class
    Dim pEnumRelClass As IEnumRelationshipClass
    Dim pRelationshipClass As IRelationshipClass
    
    Set pEnumRelClass = pFeature.Class.RelationshipClasses(esriRelRoleAny)
    Set pRelationshipClass = pEnumRelClass.Next
  
    If pRelationshipClass Is Nothing Then
        MsgBox "     No Relationship Class created for this layer." & _
        vbNewLine & "Please create a Relationship Class in ArcCatalog", _
        vbOKOnly, "Relationship Class Missing"
        Exit Sub
    End If
   
    'Loop through the features
    Do
        'Create a set that contains this object
        Dim pFeatSet As ISet
        Set pFeatSet = New esriSystem.Set
        pFeatSet.Add pFeature
        
        'Now loop through using GetObjectsRelatedToObject
        Dim i As Long
        Dim r As Long
        Dim myCount As Integer
        Dim pRelObjSet As ISet
        Dim pFieldIndex As Double
        Dim pFieldIndex2 As Double
        Dim pObjectClass As IObjectClass 'Holds the destination table
        
        Dim pRelRow As IRow
        Set pRelObjSet = pRelationshipClass.GetObjectsRelatedToObject(pFeature)
        myCount = pRelObjSet.Count
        
        Dim myArray() As String
        ReDim myArray(pRelObjSet.Count - 1)
        
        Dim myArray2() As Double
        ReDim myArray2(pRelObjSet.Count - 1)
    
        'Get the destination field
        Set pObjectClass = pRelationshipClass.DestinationClass
        pFieldIndex = pObjectClass.FindField(pFieldName)
        pFieldIndex2 = pObjectClass.FindField(pFieldName2)
        
        For i = 0 To myCount - 1
            Set pRelRow = pRelObjSet.Next
            'Get the value of the field you want to label with
            'and use it to create a label expression
            If VarType(pRelRow.value(pFieldIndex)) = vbNull Then
                myArray(i) = "-999"
            Else
                myArray(i) = pRelRow.value(pFieldIndex)
            End If
            myArray2(i) = pRelRow.value(pFieldIndex2)
        Next i
        
        'Perform a bubble sort on the data
        Dim AnyChanges As Boolean
        Dim SwapFH As Variant
        Dim SwapFH2 As Variant
        Do
            AnyChanges = False
            For i = LBound(myArray2) To UBound(myArray2) - 1
                If (myArray2(i) > myArray2(i + 1)) Then
                ' These two need to be swapped
                SwapFH = myArray2(i)
                SwapFH2 = myArray(i)
                myArray2(i) = myArray2(i + 1)
                myArray(i) = myArray(i + 1)
                myArray2(i + 1) = SwapFH
                myArray(i + 1) = SwapFH2
                AnyChanges = True
                End If
            Next i
        Loop Until Not AnyChanges 'This is the end of the bubble sort
    
        'Build label expression
        Dim strExp As String
        Dim strExp1 As String
            For i = 0 To myCount - 1
            If Not strExp = " " Then
                strExp = strExp & vbNewLine & myArray(i) & ", " & myArray2(i)
            Else
                strExp = myArray(i)
            End If
        Next i
'        Debug.Print strExp 'This is used to debug the code and data
        Call AddStick(myCount, myArray, myArray2, pFeature)
        
        'Clear the label string
        strExp = ""
        
        Set pFeature = pFeatCursor.NextFeature
    Loop Until pFeature Is Nothing
  
End Sub

Public Sub MSIExportPDF()
    Dim pMxDoc As IMxDocument
    Dim pMap As IMap
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
    Set pMxDoc = ThisDocument
    Set pPageLayout = pMxDoc.PageLayout
    Set pMxDoc.ActiveView = pPageLayout
    
'    pMxDoc.ActiveView.Refresh
    
    'Find Basemap Group and DRG layers and turn both of them on
    Set pMap = pMxDoc.FocusMap
    Set pEnum = pMap.Layers
    
    Set pLayer = pEnum.Next
    
    Do Until pLayer Is Nothing
        If pLayer.Name = "Basemaps" Then pLayer.Visible = True
        If pLayer.Name Like "DRG*" Then pLayer.Visible = True
        Set pLayer = pEnum.Next
    Loop

    'Turn off all the group layers
    Set pStdTableColl = pMap
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
    
    Set pStdTableColl = pMap
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
    
    Set pMxDoc.ActiveView = pMap
    
    pMxDoc.UpdateContents
    pMxDoc.ActiveView.Refresh
    
    
'    'Find the Layer Files metadata table
'    Set pStdTableColl = m_pMap
'    For i = 0 To pStdTableColl.StandaloneTableCount - 1
'        Set pStdTable = pStdTableColl.StandaloneTable(i)
'        If pStdTable.Name = "ODGSLAYERS" Then
'            Set pTable = pStdTable
'            lngFldLayerName = pTable.FindField("LAYERNAME")
'            lngFldPath = pTable.FindField("PATH")
'            lngFldGroupOrder = pTable.FindField("GROUPORDER")
'            lngFldGroupName = pTable.FindField("GROUPNAME")
'            lngFldGroupTOCOrder = pTable.FindField("GROUPTOCORDER")
'            lngFldGroupVis = pTable.FindField("GROUPVISABLE")
'        End If
'    Next i
'
'    If pStdTable Is Nothing Then
'        Exit Sub
'    End If
'
'    'Sort the Table
'    Set pTableSort = New TableSort
'    With pTableSort
'        .Fields = "GROUPTOCORDER, GROUPNAME"
'        .Ascending("GROUPTOCORDER") = True
'        .Ascending("GROUPNAME") = True
'        Set .QueryFilter = Nothing
'        Set .Table = pTable
'    End With
'    pTableSort.Sort Nothing
'
'    Set pCursor = pTableSort.Rows
'
'    'Find Unique Values in the Table
'    Set pDataStat = New DataStatistics
'    pDataStat.Field = "GROUPNAME"
'    Set pDataStat.Cursor = pCursor
'
'    Set pEnumVar = pDataStat.UniqueValues
'    value = pEnumVar.Next
'    Do Until IsEmpty(value)
'        'Now resort the table based upon the layer order in the group
'        Set pQF = New QueryFilter
'        pQF.WhereClause = "[GROUPNAME] = '" & value & "'"
'        Set pLyrCursor = pTable.Search(pQF, False)
'        Set pTblSortLyrs = New TableSort
'        With pTblSortLyrs
'            .Fields = "GROUPORDER"
'            .Ascending("GROUPORDER") = True
'            Set .QueryFilter = pQF
'            Set .Table = pTable
'        End With
'        pTblSortLyrs.Sort Nothing
'
'        'Get the newly sorted rows and create the new Group and Layers inside the Group
'        Set pLyrCursor = pTblSortLyrs.Rows
'        Set pRowLyrs = pLyrCursor.NextRow
'
'        'Create the new Group
'        lngGroupTOCOrder = pRowLyrs.value(lngFldGroupTOCOrder)
'        blnGroupVis = pRowLyrs.value(lngFldGroupVis)
'
'        Set pGrpLayer = New GroupLayer
'        pGrpLayer.Visible = blnGroupVis
'        pGrpLayer.Expanded = False
'        pGrpLayer.Name = pRowLyrs.value(lngFldGroupName)
'
'        'Add layers to the new Group
'        Do Until pRowLyrs Is Nothing
'            strLayerName = pRowLyrs.value(lngFldLayerName)
'            strPath = pRowLyrs.value(lngFldPath)
'
'            Set pODGSLyr = New ODGSLayer
'            Set pLayer = pODGSLyr.LoadLayer(strPath, strLayerName)
'            pGrpLayer.Add pLayer
'            Set pRowLyrs = pLyrCursor.NextRow
'        Loop
'
'        'Add the Group layer to the map
'        m_pMap.AddLayer pGrpLayer
'        m_pMap.MoveLayer pGrpLayer, lngGroupTOCOrder
'
''        Debug.Print "value - " & value & vbTab & "GroupVis = " & blnGroupVis
'        value = pEnumVar.Next
'    Loop

    
End Sub



