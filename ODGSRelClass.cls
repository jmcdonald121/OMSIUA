VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ODGSRelClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub

Function RelClassCursor(pLayer As ILayer, pQf As IQueryFilter, strRelClassName As String) As ISet
    Dim pFlayer As IFeatureLayer
    Dim pFCur As IFeatureCursor
    Dim pF As IFeature
    Dim pRelClassColl As IRelationshipClassCollection
    Dim pEnumRelClass As IEnumRelationshipClass
    Dim pRelClass As IRelationshipClass
    Dim pDS As IDataset
    Dim pRelObjSet As ISet
    
    'START
    Set pFlayer = pLayer
    Set pRelClassColl = pFlayer
    
    Set pFCur = pFlayer.Search(pQf, False)
    
    Set pF = pFCur.NextFeature
    
    'Get the Relationship Classes
    Set pEnumRelClass = pRelClassColl.RelationshipClasses
    Set pRelClass = pEnumRelClass.Next
  
    Do Until pRelClass Is Nothing
        Set pDS = pRelClass
        If pDS.BrowseName = strRelClassName Then
            Set pRelObjSet = pRelClass.GetObjectsRelatedToObject(pF)
        End If
        Set pRelClass = pEnumRelClass.Next
    Loop
    
    Set RelClassCursor = pRelObjSet

End Function
