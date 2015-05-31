Dim WithEvents m_visWins As Visio.Windows
Public WithEvents pg As Visio.Page

Private Sub Document_BeforeSelectionDelete(ByVal sel As IVSelection)

    Dim TheShapes As Shapes
    Dim deletedShape As Visio.Shape
    Set TheShapes = ActivePage.Shapes
    
    ' Unfortunately, no easy way to get notified AFTER the shape is deleted, to update the container Criteria
    ' So instead, set all shape Criteria values to 0, and trig a global refresh of all containers in the page.
    For Each deletedShape In sel
        If deletedShape.CellExists("Prop.Cost", 0) Then
            deletedShape.Cells("Prop.Cost") = 0
        End If
        
        If deletedShape.CellExists("Prop.mass", 0) Then
            deletedShape.Cells("Prop.mass") = 0#
        End If
        
        If deletedShape.CellExists("Prop.volume", 0) Then
            deletedShape.Cells("Prop.volume") = 0#
        End If
    Next
    
    ' Anything on the page has been manually deleted: refresh all containers.
    ' Parse all shapes in the page, and if it is a container, update it.
    For Each ThisShape In TheShapes
        
        If (ThisShape.CellExists("User.msvStructureType", 0)) Then
            ' Filter Containers only
            If (InStr(ThisShape.Cells("User.msvStructureType").Formula, "Container") <> 0) Then
                     UpdateContainerCriterias ThisShape
            End If
        End If
    Next

End Sub

' Use this hook called automatically on document opening to initialize our variables
Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
  Set m_visWins = Visio.Application.Windows
  Set pg = ActivePage
  
  'Refresh all container Criterias
    Dim TheShapes As Shapes
    Set TheShapes = ActivePage.Shapes
    
    ' Anything on the page has been manually deleted: refresh all containers.
    ' Parse all shapes in the page, and if it is a container, update it.
    For Each ThisShape In TheShapes
        
        If (ThisShape.CellExists("User.msvStructureType", 0)) Then
            ' Filter Containers only
            If (InStr(ThisShape.Cells("User.msvStructureType").Formula, "Container") <> 0) Then
                     UpdateContainerCriterias ThisShape
            End If
        End If
    Next
End Sub
' Since the hook relates to Page events, need to catch this page selection change event to update our local page variable
Private Sub m_visWins_WindowTurnedToPage(ByVal visWin As IVWindow)
    Set pg = ActivePage
End Sub

Private Sub pg_ContainerRelationshipAdded(ByVal ShapePair As IVRelatedShapePairEvent)
    Set modifiedContainer = ActivePage.Shapes.ItemFromID(ShapePair.FromShapeID)
    UpdateContainerCriterias modifiedContainer
End Sub


Private Sub pg_ContainerRelationshipDeleted(ByVal ShapePair As IVRelatedShapePairEvent)
    Set modifiedContainer = ActivePage.Shapes.ItemFromID(ShapePair.FromShapeID)
    UpdateContainerCriterias modifiedContainer
End Sub

' Compute and update a container's Criteria
Sub UpdateContainerCriterias(ByVal shp As Shape)
    
    Dim retShp As Shape
    Dim ShapeNames As String
    Dim totalVolume As Double
    Dim totalCost As Integer
    Dim totalMass As Double

    totalVolume = 0#
    totalCost = 0
    totalMass = 0#

    ' Compute container Criterias from contained elements's Criterias
    For Each memberID In shp.ContainerProperties.GetMemberShapes(visContainerFlagsDefault)
        Set retShp = ActivePage.Shapes.ItemFromID(memberID)
        Dim isCriteria As Integer

        isCriteria = InStr(retShp.Name, "Criteria_")
        
        ' Take into account member shape in the computation only if it's not one of the Criteria boxes
        If (isCriteria = 0) Then
            
            If retShp.CellExists("Prop.volume", 0) Then
                If retShp.Cells("Prop.volume").Formula <> "" Then
                    totalVolume = totalVolume + retShp.Cells("Prop.volume").Formula
                Else
                    MsgBox "Empty volume value on box " + retShp.Name
                End If
            Else
                MsgBox "Missing volume property on box " + retShp.Name
            End If
            
            If retShp.CellExists("Prop.Cost", 0) Then
                If retShp.Cells("Prop.Cost").Formula <> "" Then
                    totalCost = totalCost + retShp.Cells("Prop.Cost").Formula
                Else
                    MsgBox "Empty Cost value on box " + retShp.Name
                End If
            Else
                MsgBox "Missing Cost property on box " + retShp.Name
            End If
            
            If retShp.CellExists("Prop.mass", 0) Then
                If retShp.Cells("Prop.mass").Formula <> "" Then
                    totalMass = totalMass + retShp.Cells("Prop.mass").Formula
                Else
                    MsgBox "Empty Mass value on box " + retShp.Name
                End If
            Else
                MsgBox "Missing Mass property on box " + retShp.Name
            End If
        End If
    Next

    ' Update container Criterias data boxes
    For Each memberID In shp.ContainerProperties.GetMemberShapes(visContainerFlagsDefault)
        Set retShp = ActivePage.Shapes.ItemFromID(memberID)
        Dim isCriteriaVolume As Integer
        Dim isCriteriaCost As Integer
        Dim isCriteriaMass As Integer

        isCriteriaVolume = InStr(retShp.Name, "Criteria_Volume")
        isCriteriaCost = InStr(retShp.Name, "Criteria_Cost")
        isCriteriaMass = InStr(retShp.Name, "Criteria_Mass")

        If (isCriteriaVolume <> 0) Then
            retShp.Cells("Prop.criteriavalue") = totalVolume
        ElseIf (isCriteriaCost <> 0) Then
            retShp.Cells("Prop.criteriavalue") = totalCost
        ElseIf (isCriteriaMass <> 0) Then
            retShp.Cells("Prop.criteriavalue") = totalMass
        End If
    Next
End Sub














