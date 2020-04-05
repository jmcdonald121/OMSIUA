Attribute VB_Name = "ExportActiveView"
Option Explicit
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Sub ExportActiveView(strLayerName As String)

  Dim pMxDoc As IMxDocument
  Dim pActiveView As IActiveView
  Dim pExport As IExport
  Dim iPrevOutputImageQuality As Long
  Dim pOutputRasterSettings As IOutputRasterSettings
  Dim pPixelBoundsEnv As IEnvelope
  Dim exportRECT As tagRECT
  Dim DisplayBounds As tagRECT
  Dim pDisplayTransformation As IDisplayTransformation
  Dim pPageLayout As IPageLayout
  Dim pMapExtEnv As IEnvelope
  Dim hdc As Long
  Dim tmpDC As Long
  Dim sNameRoot As String
  Dim sOutputDir As String
  Dim iOutputResolution As Long
  Dim iScreenResolution As Long
  Dim bContinue As Boolean
  Dim msg As String
  Dim pTrackCancel As ITrackCancel
  Dim pGraphicsExtentEnv As IEnvelope
  Dim bClipToGraphicsExtent As Boolean
  Dim pUnitConvertor As IUnitConverter
  
  Set pMxDoc = Application.Document
  Set pActiveView = pMxDoc.ActiveView
  Set pTrackCancel = New CancelTracker
  
  'Create an ExportPDF object and QI the pExport interface pointer onto it.
  ' To export to a format other than PDF, simply create a different CoClass here
  Set pExport = New ExportPDF
  'assign a resolution for the export in dpi
  iOutputResolution = 1200
  'assign True or False to determin is export image will be clipped to the graphic extent of layout elements.
  'this value is ignored for data view exports
  bClipToGraphicsExtent = False
  
  
  Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
  iPrevOutputImageQuality = pOutputRasterSettings.ResampleRatio
  ' Output Image Quality of the export.  The value here will only be used if the export
  '  object is a format that allows setting of Output Image Quality, i.e. a vector exporter.
  '  The value assigned to ResampleRatio should be in the range 1 to 5.
  '  1 corresponds to "Best", 5 corresponds to "Fast"
  If TypeOf pExport Is IExportImage Then
  'always set the output quality of the display to 1 for image export formats
    SetOutputQuality pActiveView, 1
  ElseIf TypeOf pExport Is IOutputRasterSettings Then
  ' for vector formats, assign a ResampleRatio to control drawing of raster layers at export time
    Set pOutputRasterSettings = pExport
    pOutputRasterSettings.ResampleRatio = 1
'    Set pOutputRasterSettings = Nothing
  End If
  
  'assign the output path and filename.  We can use the Filter property of the export object to
  ' automatically assign the proper extension to the file.
  sOutputDir = "D:\OMSIUA\TMP\"
  sNameRoot = strLayerName
'  sNameRoot = strLayerName & "_" & Left(ThisDocument.Title, Len(ThisDocument.Title) - 4)
  pExport.ExportFileName = sOutputDir & sNameRoot & "." & Right(Split(pExport.Filter, "|")(1), _
                           Len(Split(pExport.Filter, "|")(1)) - 2)
  tmpDC = GetDC(0)
  iScreenResolution = GetDeviceCaps(tmpDC, 88) '88 is the win32 const for Logical pixels/inch in X)
  ReleaseDC 0, tmpDC
  pExport.Resolution = iOutputResolution
  
  If TypeOf pActiveView Is IPageLayout Then
    DisplayBounds = pActiveView.ExportFrame
    Set pMapExtEnv = pGraphicsExtentEnv
  Else
    Set pDisplayTransformation = pActiveView.ScreenDisplay.DisplayTransformation
    DisplayBounds.Left = 0
    DisplayBounds.Top = 0
    DisplayBounds.Right = pDisplayTransformation.DeviceFrame.Right
    DisplayBounds.bottom = pDisplayTransformation.DeviceFrame.bottom
    Set pMapExtEnv = New Envelope
    Set pMapExtEnv = pDisplayTransformation.FittedBounds
  End If
  
  Set pPixelBoundsEnv = New Envelope
  If bClipToGraphicsExtent And (TypeOf pActiveView Is IPageLayout) Then
    Set pGraphicsExtentEnv = GetGraphicsExtent(pActiveView)
    Set pPageLayout = pActiveView
    Set pUnitConvertor = New UnitConverter
    'assign the x and y values representing the clipped area to the PixelBounds envelope
    pPixelBoundsEnv.XMin = 0
    pPixelBoundsEnv.YMin = 0
    pPixelBoundsEnv.XMax = pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.XMax, pPageLayout.Page.Units, esriInches) * pExport.Resolution _
                          - pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.XMin, pPageLayout.Page.Units, esriInches) * pExport.Resolution
    pPixelBoundsEnv.YMax = pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.YMax, pPageLayout.Page.Units, esriInches) * pExport.Resolution _
                          - pUnitConvertor.ConvertUnits(pGraphicsExtentEnv.YMin, pPageLayout.Page.Units, esriInches) * pExport.Resolution
                          
    'assign the x and y values representing the clipped export extent to the exportRECT
    With exportRECT
      .bottom = Fix(pPixelBoundsEnv.YMax) + 1
      .Left = Fix(pPixelBoundsEnv.XMin)
      .Top = Fix(pPixelBoundsEnv.YMin)
      .Right = Fix(pPixelBoundsEnv.XMax) + 1
    End With
    
    Set pMapExtEnv = pGraphicsExtentEnv
  Else
    'The values in the exportRECT tagRECT correspond to the width
    ' and height to export, measured in pixels with an origin in the top left corner.
    With exportRECT
      .bottom = DisplayBounds.bottom * (iOutputResolution / iScreenResolution)
      .Left = DisplayBounds.Left * (iOutputResolution / iScreenResolution)
      .Top = DisplayBounds.Top * (iOutputResolution / iScreenResolution)
      .Right = DisplayBounds.Right * (iOutputResolution / iScreenResolution)
    End With
    'populate the PixelBounds envelope with the values from exportRECT.
    ' We need to do this because the exporter object requires an envelope object
    ' instead of a tagRECT structure.
    pPixelBoundsEnv.PutCoords exportRECT.Left, exportRECT.Top, exportRECT.Right, exportRECT.bottom
  End If
  
  'Assign the envelope object to the exporter object's PixelBounds property.  The exporter object
  ' will use these dimensions when allocating memory for the export file.
  pExport.PixelBounds = pPixelBoundsEnv
  
  Set pExport.TrackCancel = pTrackCancel
  Set pExport.StepProgressor = Application.StatusBar.ProgressBar
  pTrackCancel.Reset
  pTrackCancel.CancelOnClick = False
  pTrackCancel.CancelOnKeyPress = True
  bContinue = pTrackCancel.Continue()
  
  hdc = pExport.StartExporting
    
    'Redraw the active view, rendering it to the exporter object device context instead of the app display.
  'We pass the following values:
  ' * hDC is the device context of the exporter object.
  ' * exportRECT is the tagRECT structure that describes the dimensions of the view that will be rendered.
  ' The values in exportRECT should match those held in the exporter object's PixelBounds property.
  ' * pMapExtEnv is an envelope defining the section of the original image to draw into the export object.
  ' * pTrackCancel is a reference to a CancelTracker object
  pActiveView.Output hdc, pExport.Resolution, exportRECT, pMapExtEnv, pTrackCancel
  
  bContinue = pTrackCancel.Continue()
  If bContinue Then
    msg = "Writing export file..."
    Application.StatusBar.Message(0) = msg
    pExport.FinishExporting
    pExport.Cleanup
  Else
    pExport.Cleanup
  End If
  pTrackCancel.CancelOnClick = False
  pTrackCancel.CancelOnKeyPress = True
  
  bContinue = pTrackCancel.Continue()
  If bContinue Then
    msg = "Finished exporting '" & pExport.ExportFileName & "'"
    Application.StatusBar.Message(0) = msg
  End If
  
  SetOutputQuality pActiveView, iPrevOutputImageQuality
  Set pTrackCancel = Nothing
  Set pMapExtEnv = Nothing
  Set pPixelBoundsEnv = Nothing
End Sub


Private Sub SetOutputQuality(pActiveView As IActiveView, iResampleRatio As Long)
  Dim pMap As IMap
  Dim pGraphicsContainer As IGraphicsContainer
  Dim pElement As IElement
  Dim pOutputRasterSettings As IOutputRasterSettings
  Dim pMapFrame As IMapFrame
  Dim pTmpActiveView As IActiveView
  
  If TypeOf pActiveView Is IMap Then
    Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
    pOutputRasterSettings.ResampleRatio = iResampleRatio
  ElseIf TypeOf pActiveView Is IPageLayout Then
    
    'assign ResampleRatio for PageLayout
    Set pOutputRasterSettings = pActiveView.ScreenDisplay.DisplayTransformation
    pOutputRasterSettings.ResampleRatio = iResampleRatio
    
    'and assign ResampleRatio to the Maps in the PageLayout
    Set pGraphicsContainer = pActiveView
    pGraphicsContainer.Reset
    Set pElement = pGraphicsContainer.Next
    Do While Not pElement Is Nothing
      If TypeOf pElement Is IMapFrame Then
        Set pMapFrame = pElement
        Set pTmpActiveView = pMapFrame.Map
        Set pOutputRasterSettings = pTmpActiveView.ScreenDisplay.DisplayTransformation
        pOutputRasterSettings.ResampleRatio = iResampleRatio
      End If
      DoEvents
      Set pElement = pGraphicsContainer.Next
    Loop
    Set pMap = Nothing
    Set pMapFrame = Nothing
    Set pGraphicsContainer = Nothing
    Set pTmpActiveView = Nothing
  End If
  Set pOutputRasterSettings = Nothing
  
End Sub

Function GetGraphicsExtent(pActiveView As IActiveView) As IEnvelope
  Dim pBounds As IEnvelope
  Dim pEnv As IEnvelope
  Dim pGraphicsContainer As IGraphicsContainer
  Dim pPageLayout As IPageLayout
  Dim pDisplay As IDisplay
  Dim pElement As IElement
  
  Set pBounds = New Envelope
  Set pEnv = New Envelope
  Set pPageLayout = pActiveView
  Set pDisplay = pActiveView.ScreenDisplay
  Set pGraphicsContainer = pActiveView
  pGraphicsContainer.Reset
  
  Set pElement = pGraphicsContainer.Next
  Do While Not pElement Is Nothing
    pElement.QueryBounds pDisplay, pEnv
    pBounds.Union pEnv
    DoEvents
    Set pElement = pGraphicsContainer.Next
  Loop
  
  Set GetGraphicsExtent = pBounds
  
  Set pBounds = Nothing
  Set pEnv = Nothing
  Set pGraphicsContainer = Nothing
  Set pPageLayout = Nothing
  Set pDisplay = Nothing
  Set pElement = Nothing

End Function
