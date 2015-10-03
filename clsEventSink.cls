VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements Visio.IVisEventProc

Private Function IVisEventProc_VisEventProc( _
    ByVal nEventCode As Integer, _
    ByVal pSourceObj As Object, _
    ByVal nEventID As Long, _
    ByVal nEventSeqNum As Long, _
    ByVal pSubjectObj As Object, _
    ByVal vMoreInfo As Variant) As Variant
  
    ' Custom event handler
    Select Case nEventCode
        Case visEvtCodeShapeDelete
            Debug.Print "Custom EVENT visEvtCodeShapeDelete received"
            ThisDocument.onShapeDeleted
        Case visEvtCodeShapeBeforeTextEdit
            Debug.Print "Custom EVENT visEvtCodeShapeBeforeTextEdit received"
            ThisDocument.onBeforeShapeTextEdit pSubjectObj
        Case Else
            Debug.Print "EVENT " & nEventCode & " received"
    End Select

End Function

