VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Force explicit declaration of all variables
Option Explicit

' Event notification related statements
Dim WithEvents win As Visio.Windows
Attribute win.VB_VarHelpID = -1
Public WithEvents pg As Visio.Page
Attribute pg.VB_VarHelpID = -1
Private mvsoEventHandler As clsEventSink

' internal constants
Private Const ExternalConnectorLineColor As String = "RGB(0,0,0)"
Private Const ExternalInterfaceFillColor As String = "RGB(255,255,255)"

Private Const volumeIndicatorProperty As String = "volume"
Private Const costIndicatorProperty As String = "cost"
Private Const massIndicatorProperty As String = "mass"
Private Const powerIndicatorProperty As String = "power"

Private Const connectorNbSignalsPropertyName As String = "nbSignals"
Private Const connectorUserText As String = "userName"
Private Const connectorSignalsTextPrefix As String = " ("
Private Const connectorSignalsTextSuffix As String = ")"


''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''
'''''''''' EVENT HANDLERS ''''''''''''
''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''

' Many Visio Events are available by default, but some of them need an explicit registration.
Private Sub registerCustomEvents()

    ' declare our custom event handler (implemented in our clsEventSink class)
    Dim eventHandler As clsEventSink
    Set eventHandler = New clsEventSink
    
    ' Attach our custom event handler to the appropriate event codes we wish to be notified of:
    '   - shape deleted event (triggered AFTER the shape has been deleted)
    '   - shape text is about to be edited
    ThisDocument.EventList.AddAdvise visEvtCodeShapeDelete, eventHandler, "", ""
    ThisDocument.EventList.AddAdvise visEvtCodeShapeBeforeTextEdit, eventHandler, "", ""

End Sub

' Hook called automatically on document opening
Private Sub Document_DocumentOpened(ByVal Doc As IVDocument)
    Debug.Print "----Document_DocumentOpened----"
    
    ' Register to get notified of some specific events that are not available by default
    registerCustomEvents
    
    ' Keep track of active window and page
    Set win = Visio.Application.Windows
    Set pg = ActivePage
    
    ' Refresh everything, just in case.
    refreshAllContainers
    updateAllConnectors
    updateTotals
    
    Debug.Print "----END Document_DocumentOpened----"
End Sub

' Observe shape cells modification too, to catch a manual modification of a shape's associated data,
' to be able to trig a refresh of automatic criteria boxes then
Private Sub pg_CellChanged(ByVal Cell As IVCell)

    Dim parentContainer As Shape
    
    ' IMPORTANT: filter only changes on interesting (i.e. criteria-related) cells. Since many of the shape cells are updated very often,
    ' not filtering would result in a major performance drop
    If Cell.Name = "Prop." & costIndicatorProperty Or _
        Cell.Name = "Prop." & volumeIndicatorProperty Or _
        Cell.Name = "Prop." & massIndicatorProperty Or _
        Cell.Name = "Prop." & powerIndicatorProperty Or _
        Cell.Name = "Prop." & connectorNbSignalsPropertyName Or _
        Cell.Name = "Prop." & connectorUserText Then
        
        Debug.Print "----pg_CellChanged----"
        Debug.Print , " cell " & Cell.Name & " of " & Cell.Shape.Name & " changed to " & Cell.ResultStr(visNone)
        
        ' If modified shape is inside a container, refresh this container
        Set parentContainer = getParentContainer(Cell.Shape)
        If Not parentContainer Is Nothing Then
            updateContainer parentContainer
        End If
        
        ' If a container's own data has been modified, refresh it
        If isContainer(Cell.Shape) Then
            updateContainer Cell.Shape
        End If
        
        ' If a connector data has been modified, refresh it
        ' Since an update to the connector's data may impact the interface it is
        ' potentially glued to, also trig a refresh of connected shapes
        If isConnector(Cell.Shape) Then
            updateConnector Cell.Shape
            updateConnectedShapes Cell.Shape
        End If
        
        ' Refresh totals in any case
        updateTotals
        
        Set parentContainer = Nothing
        Debug.Print "----END pg_CellChanged----"

    End If
End Sub

' Capture the event of user finishing to edit the text of a shape
Private Sub pg_ShapeExitedTextEdit(ByVal Shape As IVShape)
    Debug.Print "----pg_ShapeExitedTextEdit(" & Shape.Name & ")----"
    
    ' Keep track of new name that was just entered by the user, to later rebuild the full text
    If Shape.CellExists("Prop." & connectorUserText, 0) Then
        Shape.CellsU("Prop." & connectorUserText).FormulaU = Chr(34) & Shape.Text & Chr(34)
    End If
    
    ' Trig a refresh of this connector to rebuild the full text
    If isConnector(Shape) Then
        updateConnector Shape
    End If
    
    Debug.Print "----END pg_ShapeExitedTextEdit(" & Shape.Name & ")----"
End Sub

' Callback notified when active page changes.
' Since the other callbacks relate to Page events, need to catch this page selection change event to
' keep "pg" callback prefix pointing to the right page.
Private Sub win_WindowTurnedToPage(ByVal visWin As IVWindow)

    Debug.Print "----win_WindowTurnedToPage(" & ActivePage.Name & ")----"
    
    Set pg = ActivePage
    
    Debug.Print "----END win_WindowTurnedToPage----"
    
End Sub

' Callback notified when a shape is dropped on the page.
Private Sub pg_ShapeAdded(ByVal Shape As IVShape)

    Debug.Print "----pg_ShapeAdded(" & Shape.Name & ")----"
    
    ' If a standard visio connector was dropped, need to customize it with our data
    If isConnector(Shape) Then
        updateConnector Shape
    End If
    
    ' If an Interface shape was created, refresh it
    If isInterface(Shape) Then
        updateInterface Shape
    End If
    
    ' Refresh totals (in case the added shape already has criteria values set)
    updateTotals
    
    Debug.Print "----END pg_ShapeAdded(" & Shape.Name & ")----"
    
End Sub

' This gets called by our custom event handler AFTER a shape has been deleted from the diagram
Public Sub onShapeDeleted()

    Debug.Print "----onShapeDeleted----"
    
    ' no easy way to tell which shape has been deleted from where: just refresh everything...
    refreshAllContainers
    updateTotals
    
    Debug.Print "----END onShapeDeleted----"
    
End Sub


Public Sub onBeforeShapeTextEdit(ByVal shp As Shape)
    
    Dim signalsText As String
    
    Debug.Print "----onBeforeShapeTextEdit(" & shp.Name & ")----"
    
    If isConnector(shp) Then
        ' The user only wants to edit the name of the link : temporarily reset the text of the connector
        ' to containing the name only, not the nb of signals information: it will be re-inserted automatically
        ' when refreshing the connector upon exiting the editing
        
        'If shp.CellExists("Prop." & connectorNbSignalsPropertyName, 0) Then
            'signalsText = connectorSignalsTextPrefix & shp.CellsU("Prop." & connectorNbSignalsPropertyName) & connectorSignalsTextSuffix
            'shp.Text = Replace(shp.Text, signalsText, "")
        If shp.CellExists("Prop." & connectorUserText, 0) Then
            shp.Text = shp.CellsU("Prop." & connectorUserText).ResultStrU(VisUnitCodes.visUnitsString)
        End If
    End If

    Debug.Print "----END onBeforeShapeTextEdit(" & shp.Name & ")----"

End Sub

' common function called whenever a connector is modified (connection added or deleted, i.e. any end of the connector
' is glued or unglued from a component), as detected by pb_ConnectionsXXX callbacks
Private Sub onConnectionModified(ByVal Connects As IVConnects)

    Dim conn As Visio.Connect

    Debug.Print "----onConnectionModified----"
    
    For Each conn In Connects
    
        Debug.Print , "Processing modified connection: ToSheet=" & conn.ToSheet.Name & ", FromSheet=" & conn.FromSheet.Name
       
        ' If target of the modified connection is an Interface, update this Interface's values
        If isInterface(conn.ToSheet) Then
            updateInterface conn.ToSheet
        End If
        
        ' If source of the modified connection is an Interface, also update this Interface's values
        If isInterface(conn.FromSheet) Then
            updateInterface conn.FromSheet
        End If
        
        ' If target of the modified connection is a connector, update this connector's (e.g. color)
        If isConnector(conn.ToSheet) Then
            updateConnector conn.ToSheet
        End If
                
        ' If source of the modified connection is a connector, update this connector's (e.g. color)
        If isConnector(conn.FromSheet) Then
            updateConnector conn.FromSheet
        End If
        
    Next
    
    Debug.Print "----END onConnectionModified----"
End Sub

Private Sub pg_ConnectionsAdded(ByVal Connects As IVConnects)

    Debug.Print "----pg_ConnectionsAdded----"
    
    ' update the source and destination of the added connector
    onConnectionModified Connects
    
    Debug.Print "----END pg_ConnectionsAdded----"
    
End Sub

Private Sub pg_ConnectionsDeleted(ByVal Connects As IVConnects)

    Debug.Print "----pg_ConnectionsDeleted----"
    
    ' update the source and destination of the deleted connector
    onConnectionModified Connects
    
    Debug.Print "----END pg_ConnectionsDeleted----"
    
End Sub

' Callback notified when something is added inside a container
Private Sub pg_ContainerRelationshipAdded(ByVal ShapePair As IVRelatedShapePairEvent)
    
    Dim addedShape As Shape
    Dim impactedContainer As Shape
    
    Debug.Print "----pg_ContainerRelationshipAdded----"

    'Note: in this context, in the ShapePair input param
    ' the "From" shape is the container,
    ' the "To" shape is the shape being removed from this container.
    Set addedShape = ActivePage.Shapes.ItemFromID(ShapePair.ToShapeID)
    Set impactedContainer = ActivePage.Shapes.ItemFromID(ShapePair.FromShapeID)
    
    Debug.Print , "Shape " & addedShape.Name & " added to container " & impactedContainer.Name
    
    ' Since a shape was added to this container, refresh its values
    updateContainer impactedContainer
    
    ' Update the added shape if necessary:
    ' - if added shape is a connector, change its color to reflect the fact that it is now internal to the container
   
    If isConnector(addedShape) Then
        updateConnector addedShape
    End If
    
    Set addedShape = Nothing
    Set impactedContainer = Nothing
    
    Debug.Print "----END pg_ContainerRelationshipAdded----"

End Sub

' Callback notified when something is removed from a container
Private Sub pg_ContainerRelationshipDeleted(ByVal ShapePair As IVRelatedShapePairEvent)

    Dim removedShape As Shape
    Dim impactedContainer As Shape
    Dim newContainer As Shape
    Debug.Print "----pg_ContainerRelationshipDeleted----"
    
    'Note: in this context, in the ShapePair input param
    ' the "From" shape is the container,
    ' the "To" shape is the shape being removed from this container
    Set removedShape = ActivePage.Shapes.ItemFromID(ShapePair.ToShapeID)
    Set impactedContainer = ActivePage.Shapes.ItemFromID(ShapePair.FromShapeID)
    
    Debug.Print , "Shape " & removedShape.Name & " removed from container " & impactedContainer.Name
    
    ' Since a shape was removed from this container, refresh its values
    updateContainer impactedContainer
          
    ' Update the removed shape if necessary
    ' - if removed shape is a criteria box,
    '   - reset its value to 0 if it is now external / not in any container
    '   - if it is now in a new container, let this new container take care of updating it
    ' - if removed shape is a connector, change its color to reflect the fact that it is now external
    ' - if removed shape is an Interface, need to refresh its links counts (since some of the links that were
    '   internal may now become external)
    
    If isCriteria(removedShape) Then
        Set newContainer = getParentContainer(removedShape)
        If newContainer Is Nothing Then
            removedShape.CellsU("Prop.criteriavalue") = 0
            Debug.Print "removed indicator " & removedShape.Name & " is now external: RESET its value to 0"
        End If
    ElseIf isConnector(removedShape) Then
        updateConnector removedShape
    ElseIf isInterface(removedShape) Then
        updateInterface removedShape
    End If
    
    Set removedShape = Nothing
    Set impactedContainer = Nothing
    Set newContainer = Nothing
    Debug.Print "----END pg_ContainerRelationshipDeleted----"
    
End Sub

''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''
''''''''''''' UTILITIES ''''''''''''''
''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''
Private Function isContainer(ByVal shp As Shape) As Boolean

    isContainer = False
    If (shp.CellExists("User.msvStructureType", 0)) Then
        If (InStr(shp.CellsU("User.msvStructureType").ResultStr(visNone), "Container") <> 0) Then
            isContainer = True
        End If
    End If
    
End Function

Private Function isCriteria(ByVal shp As Shape) As Boolean
    isCriteria = shp.CellExists("Prop.criteriavalue", 0)
End Function

Private Function isInterface(ByVal shp As Shape) As Boolean
    isInterface = (shp.CellExists("Prop.interfaceInternalLinks", 0) And shp.CellExists("Prop.interfaceExternalLinks", 0))
End Function

Private Function isConnector(ByVal shp As Shape) As Boolean
   'Assume that 1D shapes are connectors
   isConnector = shp.OneD
End Function

Private Function isLink(ByVal shp As Shape) As Boolean
    isLink = shp.OneD
End Function

Private Function getParentContainer(ByVal shp As Shape) As Shape
   
    Dim containers() As Long
    Dim i As Long
    
    containers = shp.MemberOfContainers
       
    If (UBound(containers) - LBound(containers) >= 0) Then
        Set getParentContainer = shp.ContainingPage.Shapes.ItemFromID(containers(0))
    Else
        Set getParentContainer = Nothing
    End If
    
End Function

Public Function findTopLevelGroupForShape(ByVal shp As Shape) As Shape

    Dim consideredShape As Visio.Shape
    Set consideredShape = shp
    
    ' Iteratively find out if this shape is inside a group, and if this group is itself in a group, etc...
    ' until finding top level group (the parent of the top level group being the Visio Page itself)
    While (consideredShape.Parent.Type <> visTypePage)
        Set consideredShape = consideredShape.Parent
    Wend
   
   Set findTopLevelGroupForShape = consideredShape
   
   Set consideredShape = Nothing
    
End Function

' Among all possible Shape objects that may exist on a page as per the Visio objet model, only a very small subset
' are of interest in our context: to avoid having to handle many specific cases all over the code to ignore shapes
' of uninteresting types, this function filters them out and should be used wherever we iterate over all shapes on the page
Public Function filteredShapes(ByVal inputShapes As Shapes) As collection

' Shape types are:
'visTypeBitmap
'visTypeDoc
'visTypeForeignObject
'visTypeGroup
'visTypeGuide
'visTypeInk
'visTypeInval
'visTypeIsControl
'visTypeIsEmbedded
'visTypeIsLinked
'visTypeIsOLE2
'visTypeMetafile
'visTypePage
'visTypeShape

Dim collection As New collection
Dim x As Shape
Dim shp As Shape

    For Each shp In inputShapes
        Select Case shp.Type
            ' Only Page, Shapes, and Groups of Shapes are of interest to us
            Case visTypePage, visTypeShape, visTypeGroup
                collection.Add shp
            Case Else
                Debug.Print "filteredShapes: ignoring object " + shp.Name
        End Select
    Next
        
    Set filteredShapes = collection

End Function

' Same as above, but taking a list of shape Ids as input instead of a collection of Shapes
Public Function filteredShapesByIds(ByRef inputShapes() As Long) As collection

Dim collection As New collection
Dim x As Shape
Dim shpId As Variant
Dim shp As Shape

    For Each shpId In inputShapes
        Set shp = ActivePage.Shapes.ItemFromID(shpId)
        Select Case shp.Type
            Case visTypePage, visTypeShape, visTypeGroup
                collection.Add shp
            Case Else
                Debug.Print "filteredShapes: ignoring object " + shp.Name
        End Select
    Next
        
    Set filteredShapesByIds = collection
    Set shp = Nothing

End Function

Private Function countConnections(ByVal shp As Shape) As Variant

Dim returnedArray(2) As Variant
Dim sourceShape As Visio.Shape
Dim targetShape As Visio.Shape
Dim connectedShape As Visio.Shape
Dim container As Shape
Dim connectorContainer As Shape
Dim aryTargetIDs() As Long
Dim arySourceIDs() As Long
Dim allIDs() As Long
Dim i As Integer
Dim nbIncoming As Long
Dim nbOutgoing As Long
Dim topLevelGroupToSheet As Shape
Dim topLevelGroupFromSheet As Shape

    Debug.Print "----countConnections(" & shp.Name & ")----"
    
    returnedArray(0) = 0
    returnedArray(1) = 0

    arySourceIDs = shp.GluedShapes(visGluedShapesIncoming1D, "")
    aryTargetIDs = shp.GluedShapes(visGluedShapesOutgoing1D, "")

    If (UBound(arySourceIDs) = -1) Then
        nbIncoming = 0
    Else
        nbIncoming = UBound(arySourceIDs) + 1
    End If
   
    If (UBound(aryTargetIDs) = -1) Then
        nbOutgoing = 0
    Else
        nbOutgoing = UBound(aryTargetIDs) + 1
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Handle the simplest case where no links are present on this
    ' Interface
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (nbIncoming + nbOutgoing) = 0 Then
        Debug.Print , "No connections on shape " & shp.Name & ", not counting"
        countConnections = returnedArray
        Exit Function
    End If
    
    Set container = getParentContainer(shp)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Handle the simple case where this Interface object is outside any container: by definition the
    ' nb of internal links is 0, and the nb of external links is just the total of incoming and outgoing links,
    ' (times the nb of signals per link)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If container Is Nothing Then
        returnedArray(0) = 0
        returnedArray(1) = nbOutgoing + nbIncoming
        countConnections = returnedArray
        Debug.Print , "Interface " & shp.Name & " is outside any container, returning externalCount=" & returnedArray(1)
        Exit Function
    End If
   
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Then handle the more complex case where the interface is inside a container: in this case need to count
    ' which links are internal, and which links are external to the parent container.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' First concatenate arrays of incoming and outgoing visio connectors
    ReDim allIDs(nbIncoming + nbOutgoing - 1)

    For i = 0 To UBound(allIDs)
        If i <= (nbIncoming - 1) Then
            allIDs(i) = arySourceIDs(i)
        Else
            allIDs(i) = aryTargetIDs(i - UBound(arySourceIDs) - 1)
        End If
    Next

    ' Among all incoming and outgoing connector links on this Interface object, sort them as being either internal or external to the container
    For i = 0 To UBound(allIDs)
        Set connectedShape = Visio.ActivePage.Shapes.ItemFromID(allIDs(i))
        Set connectorContainer = getParentContainer(connectedShape)

        ' If connector is inside the container, count the link as internal
        If Not connectorContainer Is Nothing Then
            Debug.Print , "connector " & connectedShape.Name & " is a member of the same container (" & container.Name & " ), incrementing INTERNAL links counter"
            
            If connectedShape.CellExists("Prop." & connectorNbSignalsPropertyName, 0) Then
                ' add the number of signals carried by this link to the grand total
                returnedArray(0) = returnedArray(0) + connectedShape.CellsU("Prop." & connectorNbSignalsPropertyName)
                Debug.Print , "connector " & connectedShape.Name & " carries " & connectedShape.CellsU("Prop." & connectorNbSignalsPropertyName) & " signals"
            Else
               ' by default consider that the link carriers a single signal
               returnedArray(0) = returnedArray(0) + 1
               Debug.Print , "connector " & connectedShape.Name & " has no signals property: counting as 1 signal"
            End If
        Else
        ' If connector is outside a container, count it as external
        
            Debug.Print , "connector " & connectedShape.Name; " is NOT a member of the same container (" & container.Name & " ), counting as EXTERNAL link"
        
            If connectedShape.CellExists("Prop." & connectorNbSignalsPropertyName, 0) Then
                ' add the number of signals carried by this link to the grand total
                returnedArray(1) = returnedArray(1) + connectedShape.CellsU("Prop." & connectorNbSignalsPropertyName)
                Debug.Print , "connector " & connectedShape.Name & " carries " & connectedShape.CellsU("Prop." & connectorNbSignalsPropertyName) & " signals"
            Else
               ' by default consider that the link carriers a single signal
               returnedArray(1) = returnedArray(1) + 1
               Debug.Print , "connector " & connectedShape.Name & " has no signals property: counting as 1 signal"
            End If
        End If
    Next

    countConnections = returnedArray
    Debug.Print , "countConnections returned internalCount=" & returnedArray(0) & ", externalCount=" & returnedArray(1)

    Set sourceShape = Nothing
    Set targetShape = Nothing
    Set connectedShape = Nothing
    Set container = Nothing
    Set connectorContainer = Nothing
    Set topLevelGroupToSheet = Nothing
    Set topLevelGroupFromSheet = Nothing
    
    Debug.Print "----END countConnections(" & shp.Name & ")----"
End Function

''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''
'''''''''' CORE FUNCTIONS ''''''''''''
''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''

Private Sub updateConnector(ByVal shp As Shape)
    
    Debug.Print "----updateConnector(" & shp.Name & ")----"

        If Application.IsUndoingOrRedoing Then
            Debug.Print "UNDO context: skipping updateConnector"
            Exit Sub
        End If

    Dim containerShp As Shape
    Dim newPropRow As Integer
    Dim conn As Visio.Connect
    Dim toShape As Shape
    Dim topLevelGroup As Shape
    Dim signalsText As String
    Dim userText As String
    
    Set containerShp = getParentContainer(shp)
    
    ' Check if connector has our required custom shape data, and if not (e.g. connector was created by user from the menu) add it
    If Not shp.CellExists("Prop." & connectorNbSignalsPropertyName, 0) Then
        Debug.Print , "Connector " & shp.Name & " does not have custom property Prop." & connectorNbSignalsPropertyName & ", adding it"
        newPropRow = shp.AddNamedRow(visSectionProp, connectorNbSignalsPropertyName, visTagDefault)
        ' By default, consider that a link carries a single signal
        shp.CellsU("Prop." & connectorNbSignalsPropertyName).FormulaU = "1"
    End If
    
    If Not shp.CellExists("Prop." & connectorUserText, 0) Then
        Debug.Print , "Connector " & shp.Name & " does not have custom property Prop." & connectorUserText & ", adding it"
        newPropRow = shp.AddNamedRow(visSectionProp, connectorUserText, visTagDefault)
        ' By default, just copy whatever was the current link name/text at that time
        shp.CellsU("Prop." & connectorUserText).FormulaU = Chr(34) & shp.Text & Chr(34)
    End If
    
    ' Refresh the connector text, taking into account the custom data on the nb of signals carried by this connector
    If shp.CellExists("Prop." & connectorNbSignalsPropertyName, 0) And shp.CellExists("Prop." & connectorUserText, 0) Then
        signalsText = connectorSignalsTextPrefix & shp.CellsU("Prop." & connectorNbSignalsPropertyName) & connectorSignalsTextSuffix
        userText = shp.CellsU("Prop." & connectorUserText).ResultStrU(VisUnitCodes.visUnitsString)
        shp.Text = userText & signalsText
        Debug.Print , "Updated connector text to " & shp.Text
    End If

    ' If connector is inside a container, assign it the fill color of its parent container
    If Not containerShp Is Nothing Then
    
       If shp.CellExists("LineColor", 0) Then
            shp.CellsU("LineColor").FormulaU = containerShp.Name & "!FillForegnd"
        Else
            MsgBox "ERROR: No LineColor attribute on shape " & shp.Name
        End If

        ' Also adjust its text color accordingly
        shp.CellsSRC(visSectionCharacter, 0, visCharacterColor).FormulaU = containerShp.Name & "!FillForegnd"
        
    Else
    
    ' If connector is outside a container, assign it the default color for external connectors
        If shp.CellExists("LineColor", 0) Then
            shp.CellsU("LineColor").FormulaU = ExternalConnectorLineColor
        Else
            MsgBox "ERROR: No LineColor attribute on shape " & shp.Name
        End If
        
        ' Also adjust its text color accordingly
        shp.CellsSRC(visSectionCharacter, 0, visCharacterColor).FormulaU = ExternalConnectorLineColor
        
    End If
    
    Set containerShp = Nothing
    Set conn = Nothing
    Set toShape = Nothing
    Set topLevelGroup = Nothing
    
    Debug.Print "----END updateConnector(" & shp.Name & ")----"

End Sub

Private Sub updateConnectedShapes(ByVal connector As Shape)

    Dim conn As Connect
   
        If Application.IsUndoingOrRedoing Then
            Debug.Print "UNDO context: skipping updateConnectedShapes"
            Exit Sub
        End If
   
   
    Debug.Print "----updateConnectedShapes(" & connector.Name & ")----"
    
    ' Sanity check
    If Not isConnector(connector) Then
        Debug.Print , "ERROR: input shape is not a connector, exiting"
        Exit Sub
    End If

    For Each conn In connector.Connects
        If isInterface(conn.ToSheet) Then
            Debug.Print , "refreshing connected shape" & conn.ToSheet.Name
            updateInterface conn.ToSheet
        End If
    Next

    Debug.Print "----END updateConnectedShapes(" & connector.Name & ")----"
End Sub

Private Sub updateInterface(ByVal shp As Shape)
        Dim connectionCounts As Variant
        Dim containerShp As Shape
       
        If Application.IsUndoingOrRedoing Then
            Debug.Print "UNDO context: skipping updateInterface"
            Exit Sub
        End If

On Error GoTo ErrTrap
    
       
        Debug.Print "----updateInterface(" & shp.Name & ")----"
        
        connectionCounts = countConnections(shp)
        shp.CellsU("Prop.interfaceInternalLinks") = connectionCounts(0)
        shp.CellsU("Prop.interfaceExternalLinks") = connectionCounts(1)
        
        Set containerShp = getParentContainer(shp)
        
        If Not containerShp Is Nothing Then
            Debug.Print , "updateInterface: Interface (" & shp.Name & ") is inside a container (" & containerShp.Name & "), updating it"
                        
            ' Align the Interface border/line color to the color of the Container header
            If shp.CellExists("FillForegnd", 0) Then
                shp.CellsU("FillForegnd").FormulaU = "TINT(" & containerShp.Name & "!FillForegnd" & ", 50)"
            Else
                MsgBox "ERROR: No LineColor attribute on shape " & shp.Name
            End If
        Else
            Debug.Print , "updateInterface: Interface (" & shp.Name & ") is not inside a container"
            
            ' Align the Interface border/line color to the default color
            If shp.CellExists("FillForegnd", 0) Then
                shp.CellsU("FillForegnd").FormulaU = ExternalInterfaceFillColor
            Else
                MsgBox "ERROR: No LineColor attribute on shape " & shp.Name
            End If
        End If
        
        Set containerShp = Nothing
        Debug.Print "----END updateInterface(" & shp.Name & ")----"
       
 
Exit Sub
ErrTrap:

    MsgBox "error number " & Err.Number & " occured " & Err.Description
    Exit Sub  ' or resume next to continue
       
       
End Sub


' Compute and update a container's Criteria box values
Public Sub updateContainer(ByVal updatedContainer As Shape)
    
    Debug.Print "----updateContainer(" & updatedContainer.Name & ")----"
    
    
         If Application.IsUndoingOrRedoing Then
            Debug.Print "UNDO context: skipping updateContainer"
            Exit Sub
        End If
    
    Dim shp As Shape
    Dim memberID As Variant
    Dim ShapeNames As String
    Dim totalVolume As Double
    Dim totalCost As Integer
    Dim totalMass As Double
    Dim totalPower As Double
    totalVolume = 0#
    totalCost = 0
    totalMass = 0#

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Scan all regular shapes inside the container (excluding special shapes) and perform computations
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'For Each memberID In filteredShapes(updatedContainer.ContainerProperties.GetMemberShapes(visContainerFlagsDefault))
    For Each shp In filteredShapesByIds(updatedContainer.ContainerProperties.GetMemberShapes(visContainerFlagsDefault))
        'Set shp = ActivePage.Shapes.ItemFromID(memberID)
        
        If (isCriteria(shp) = False And isLink(shp) = False And isInterface(shp) = False) Then
        
            ' Take VOLUME value into account in container's total
            If shp.CellExists("Prop." & volumeIndicatorProperty, 0) Then
                If shp.CellsU("Prop." & volumeIndicatorProperty).ResultStr(visNone) <> "" Then
                    totalVolume = totalVolume + shp.CellsU("Prop." & volumeIndicatorProperty).Result(visNone)
                Else
                    Debug.Print "Empty volume value on shape " + shp.Name
                End If
            Else
                Debug.Print "Missing volume property on box " + shp.Name
            End If
            
            ' Take COST value into account in container's total
            If shp.CellExists("Prop." & costIndicatorProperty, 0) Then
                If shp.CellsU("Prop." & costIndicatorProperty).ResultStr(visNone) <> "" Then
                    totalCost = totalCost + shp.CellsU("Prop." & costIndicatorProperty).Result(visNone)
                Else
                    Debug.Print "Empty Cost value on shape " + shp.Name
                End If
            Else
                Debug.Print "Missing Cost property on box " + shp.Name
            End If
            
            ' Take MASS value into account in container's total
            If shp.CellExists("Prop." & massIndicatorProperty, 0) Then
                If shp.CellsU("Prop." & massIndicatorProperty).ResultStr(visNone) <> "" Then
                    totalMass = totalMass + shp.CellsU("Prop." & massIndicatorProperty).Result(visNone)
                Else
                    Debug.Print "Empty Mass value on shape " + shp.Name
                End If
            Else
                Debug.Print "Missing Mass property on box " + shp.Name
            End If
            
            ' Take POWER value into account in container's total
            If shp.CellExists("Prop." & powerIndicatorProperty, 0) Then
                If shp.CellsU("Prop." & powerIndicatorProperty).ResultStr(visNone) <> "" Then
                    totalPower = totalPower + shp.CellsU("Prop." & powerIndicatorProperty).Result(visNone)
                Else
                    Debug.Print "Empty Power value on shape " + shp.Name
                End If
            Else
                Debug.Print "Missing Power property on box " + shp.Name
            End If
        End If
    
    Next
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Also add criteria values of the container itself
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Take container's VOLUME value into account
    If updatedContainer.CellExists("Prop." & volumeIndicatorProperty, 0) Then
        If updatedContainer.CellsU("Prop." & volumeIndicatorProperty).ResultStr(visNone) <> "" Then
            totalVolume = totalVolume + updatedContainer.CellsU("Prop." & volumeIndicatorProperty).Result(visNone)
        Else
            Debug.Print "Empty volume value on shape " + updatedContainer.Name
        End If
    Else
        Debug.Print "Missing volume property on box " + updatedContainer.Name
    End If
    
    ' Take container's COST value into account
    If updatedContainer.CellExists("Prop." & costIndicatorProperty, 0) Then
        If updatedContainer.CellsU("Prop." & costIndicatorProperty).ResultStr(visNone) <> "" Then
            totalCost = totalCost + updatedContainer.CellsU("Prop." & costIndicatorProperty).Result(visNone)
        Else
            Debug.Print "Empty Cost value on shape " + updatedContainer.Name
        End If
    Else
        Debug.Print "Missing Cost property on box " + updatedContainer.Name
    End If
    
    ' Take container's MASS value into account
    If updatedContainer.CellExists("Prop." & massIndicatorProperty, 0) Then
        If updatedContainer.CellsU("Prop." & massIndicatorProperty).ResultStr(visNone) <> "" Then
            totalMass = totalMass + updatedContainer.CellsU("Prop." & massIndicatorProperty).Result(visNone)
        Else
            Debug.Print "Empty Mass value on shape " + updatedContainer.Name
        End If
    Else
        Debug.Print "Missing Mass property on box " + updatedContainer.Name
    End If
    
    ' Take container POWER value into account
    If updatedContainer.CellExists("Prop." & powerIndicatorProperty, 0) Then
        If updatedContainer.CellsU("Prop." & powerIndicatorProperty).ResultStr(visNone) <> "" Then
            totalPower = totalPower + updatedContainer.CellsU("Prop." & powerIndicatorProperty).Result(visNone)
        Else
            Debug.Print "Empty Power value on shape " + updatedContainer.Name
        End If
    Else
        Debug.Print "Missing Power property on box " + updatedContainer.Name
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Update criterias & Interface info data boxes
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'For Each memberID In filteredShapes(updatedContainer.ContainerProperties.GetMemberShapes(visContainerFlagsDefault))
    For Each shp In filteredShapesByIds(updatedContainer.ContainerProperties.GetMemberShapes(visContainerFlagsDefault))
        'Set shp = ActivePage.Shapes.ItemFromID(memberID)
        
        Debug.Print , "Parsing shape: " & shp.Name
        
        If isCriteria(shp) Then
            Select Case shp.CellsU("Prop.criteriavalue.Label").ResultStr(visNone)
                Case "Volume"
                    shp.CellsU("Prop.criteriavalue") = totalVolume
                Case "Cost"
                    shp.CellsU("Prop.criteriavalue") = totalCost
                Case "Mass"
                    shp.CellsU("Prop.criteriavalue") = totalMass
                Case "Power"
                    shp.CellsU("Prop.criteriavalue") = totalPower
            End Select
        End If
        
        If isInterface(shp) Then
            updateInterface shp
        End If
        
    Next
    
    Set shp = Nothing
    Debug.Print "----END updateContainer(" & updatedContainer.Name & ")----"
End Sub

Private Sub refreshAllContainers()
    Dim shp As Shape
    
    ' Anything on the page has been manually deleted: refresh all containers.
    ' Parse all shapes in the page, and if it is a container, update it.
    For Each shp In filteredShapes(ActivePage.Shapes)
        If isContainer(shp) Then
              updateContainer shp
        End If
    Next
    
    Set shp = Nothing
End Sub

' Compute and update global criteria values
Public Sub updateTotals()
    
    Debug.Print "----updateTotals----"
    
    Dim shp As Shape
    Dim parentContainer As Shape
    Dim memberID As Variant
    Dim ShapeNames As String
    Dim totalVolume As Double
    Dim totalCost As Integer
    Dim totalMass As Double
    Dim totalPower As Double
    totalVolume = 0#
    totalCost = 0
    totalMass = 0#

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Scan all regular shapes on the page (excluding special shapes) and perform computations
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each shp In filteredShapes(ActivePage.Shapes)
           
        If (isCriteria(shp) = False And isLink(shp) = False And isInterface(shp) = False) Then
        
            ' Take VOLUME value into account in container's total
            If shp.CellExists("Prop." & volumeIndicatorProperty, 0) Then
                If shp.CellsU("Prop." & volumeIndicatorProperty).ResultStr(visNone) <> "" Then
                    totalVolume = totalVolume + shp.CellsU("Prop." & volumeIndicatorProperty).Result(visNone)
                Else
                    Debug.Print "Empty volume value on shape " + shp.Name
                End If
            Else
                Debug.Print "Missing volume property on box " + shp.Name
            End If
            
            ' Take COST value into account in container's total
            If shp.CellExists("Prop." & costIndicatorProperty, 0) Then
                If shp.CellsU("Prop." & costIndicatorProperty).ResultStr(visNone) <> "" Then
                    totalCost = totalCost + shp.CellsU("Prop." & costIndicatorProperty).Result(visNone)
                Else
                    Debug.Print "Empty Cost value on shape " + shp.Name
                End If
            Else
                Debug.Print "Missing Cost property on box " + shp.Name
            End If
            
            ' Take MASS value into account in container's total
            If shp.CellExists("Prop." & massIndicatorProperty, 0) Then
                If shp.CellsU("Prop." & massIndicatorProperty).ResultStr(visNone) <> "" Then
                    totalMass = totalMass + shp.CellsU("Prop." & massIndicatorProperty).Result(visNone)
                Else
                    Debug.Print "Empty Mass value on shape " + shp.Name
                End If
            Else
                Debug.Print "Missing Mass property on box " + shp.Name
            End If
            
            ' Take POWER value into account in container's total
            If shp.CellExists("Prop." & powerIndicatorProperty, 0) Then
                If shp.CellsU("Prop." & powerIndicatorProperty).ResultStr(visNone) <> "" Then
                    totalPower = totalPower + shp.CellsU("Prop." & powerIndicatorProperty).Result(visNone)
                Else
                    Debug.Print "Empty Power value on shape " + shp.Name
                End If
            Else
                Debug.Print "Missing Power property on box " + shp.Name
            End If
        End If
    
    Next
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Update global criterias info data boxes
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For Each shp In filteredShapes(ActivePage.Shapes)
            
        Set parentContainer = getParentContainer(shp)
        Debug.Print , "Parsing page-level shape: " & shp.Name
        
        If isCriteria(shp) And parentContainer Is Nothing Then
            Debug.Print "Updating global criteria: " & shp.Name
            Select Case shp.CellsU("Prop.criteriavalue.Label").ResultStr(visNone)
                Case "Volume"
                    shp.CellsU("Prop.criteriavalue") = totalVolume
                Case "Cost"
                    shp.CellsU("Prop.criteriavalue") = totalCost
                Case "Mass"
                    shp.CellsU("Prop.criteriavalue") = totalMass
                Case "Power"
                    shp.CellsU("Prop.criteriavalue") = totalPower
            End Select
        End If
        
    Next

    Set shp = Nothing
    Set parentContainer = Nothing

    Debug.Print "----END updateTotals----"
End Sub

' Refresh all connectors
Public Sub updateAllConnectors()
    
    Debug.Print "----updateAllConnectors----"
    
    Dim shp As Shape

    For Each shp In filteredShapes(ActivePage.Shapes)
        If isConnector(shp) Then
            updateConnector shp
        End If
    Next
    
    Set shp = Nothing
   
    Debug.Print "----END updateAllConnectors----"
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''' DEBUG  FUNCTIONS  '''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' DEBUG/Utility function to easily inspect the attributes of a selected shape
Public Sub DEBUG_INSPECT_SHAPE()
    Dim vsoShape As Visio.Shape
    Dim selected As Visio.Selection
    Set selected = ActiveWindow.Selection
    
    ' Without this line, when selecting a sub-shape inside a group, the selection returns Nothing.
    selected.IterationMode = Visio.VisSelectMode.visSelModeSkipSuper
    ' place a breakpoint on this line, then add a spy on the variable.
    Set vsoShape = selected.PrimaryItem

End Sub

