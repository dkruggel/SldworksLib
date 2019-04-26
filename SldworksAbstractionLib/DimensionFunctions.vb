Imports SldWorks, System.Runtime.CompilerServices

Namespace DimensionAbstractions
    Public Module DimensionAbstractions

        <Extension>
        Public Function DriveDimensionValue(ByVal modelObj As ModelDoc2, ByVal dimName As String, ByVal dimVal As Double, Optional ByVal toleranceVal As Double = 0.0) As Integer
            Dim retVal As Integer '-1 = failure, 0 = dimension not selected, 1 = success
            Dim ext As ModelDocExtension = modelObj.Extension
            Dim selMgr As SelectionMgr = modelObj.SelectionManager
            Dim dispDim As DisplayDimension, dimObj As Dimension

            Try
                Dim bool As Boolean = ext.SelectByID2(dimName, "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                If bool Then
                    dispDim = selMgr.GetSelectedObject6(1, -1)
                    dimObj = dispDim.GetDimension
                    dimObj.SystemValue = dimVal
                    retVal = 1
                    modelObj.ClearSelection2(True)
                Else
                    retVal = 0
                End If
            Catch ex As Exception
                retVal = -1
                modelObj.ClearSelection2(True)
            End Try

            Return retVal
        End Function

        <Extension>
        Public Function CheckDimensionValue(ByVal modelObj As ModelDoc2, ByVal dimName As String) As Double
            Dim retVal As Double = 0.0
            Dim ext As ModelDocExtension = modelObj.Extension
            Dim selMgr As SelectionMgr = modelObj.SelectionManager
            Dim dispDim As DisplayDimension, dimObj As Dimension

            Try
                Dim bool As Boolean = ext.SelectByID2(dimName, "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                If bool Then
                    dispDim = selMgr.GetSelectedObject6(1, -1)
                    dimObj = dispDim.GetDimension
                    retVal = dimObj.SystemValue
                End If
            Catch ex As Exception
                retVal = -1.0
                modelObj.ClearSelection2(True)
            End Try

            Return retVal
        End Function

        <Extension>
        Public Function DriveDimensionPlacement(ByVal modelObj As ModelDoc2, ByVal dimName As String, ByVal xPos As Double, ByVal yPos As Double, ByVal cntr As Boolean, ByVal hide As Boolean) As Integer
            Dim bool As Boolean, dispDim As DisplayDimension, annObj As Annotation
            Dim ext As ModelDocExtension = modelObj.Extension
            Dim selMgr As SelectionMgr = modelObj.SelectionManager

            Try
                bool = ext.SelectByID2(dimName, "DIMENSION", 0, 0, 0, False, 0, Nothing, 0)
                If Not bool = False Then
                    dispDim = selMgr.GetSelectedObject6(1, -1)
                    annObj = dispDim.GetAnnotation
                    If hide Then
                        modelObj.HideDimension()
                        modelObj.ClearSelection2(True)
                    Else
                        bool = annObj.SetPosition2(xPos, yPos, 0)
                        dispDim.CenterText = cntr
                    End If
                    modelObj.GraphicsRedraw2()
                End If
            Catch ex As Exception
                Return -1
            End Try

            Return 1
        End Function

    End Module
End Namespace

Namespace NoteAbstractions
    Public Module NoteAbstractions

        <Extension>
        Public Function DriveNoteText(ByVal modelObj As ModelDoc2, ByVal noteName As String, ByVal noteText As String, Optional ByVal sizeInPoints As Integer = 0, Optional ByVal xPos As Double = 0.0, Optional ByVal yPos As Double = 0.0) As Integer
            Dim bool As Boolean, noteObj As Note, annObj As Annotation
            Dim ext As ModelDocExtension = modelObj.Extension
            Dim selMgr As SelectionMgr = modelObj.SelectionManager

            Try
                bool = ext.SelectByID2(noteName, "NOTE", 0, 0, 0, False, 0, Nothing, 0)
                If bool Then
                    noteObj = selMgr.GetSelectedObject6(1, -1)
                    If Not sizeInPoints = 0 Then
                        noteObj.SetHeightInPoints(sizeInPoints)
                    End If
                    noteObj.SetText(noteText)
                    If Not xPos = 0 Then
                        annObj = noteObj.GetAnnotation
                        bool = annObj.SetPosition2(xPos, yPos, 0)
                    End If
                End If
            Catch ex As Exception
                Return -1
            End Try

            Return 0
        End Function

    End Module
End Namespace

Namespace ViewAbstractions
    Public Module ViewAbstractions

        <Extension>
        Public Function DriveViewPosition(ByVal modelObj As ModelDoc2, ByVal viewName As String, ByVal xPos As Double, ByVal yPos As Double) As Integer
            Dim bool As Boolean, tempView As SldWorks.View
            Dim ext As ModelDocExtension = modelObj.Extension
            Dim selMgr As SelectionMgr = modelObj.SelectionManager

            Try
                bool = ext.SelectByID2(viewName, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
                If bool Then
                    tempView = selMgr.GetSelectedObject6(1, -1)
                    tempView.Position = {xPos, yPos}
                End If
            Catch ex As Exception
                Return -1
            End Try

            Return 1
        End Function

    End Module
End Namespace

Namespace FeatureAbstractions
    Public Module FeatureAbstractions

        <Extension>
        Public Function FeatSuppressed(ByVal modelObj As ModelDoc2, ByVal featName As String) As Boolean
            Dim bool As Boolean, tempFeature As Feature
            Dim ext As ModelDocExtension = modelObj.Extension
            Dim selMgr As SelectionMgr = modelObj.SelectionManager

            bool = ext.SelectByID2(featName, "BODYFEATURE", 0, 0, 0, False, 0, Nothing, 0)
            tempFeature = selMgr.GetSelectedObject6(1, -1)
            Return tempFeature.IsSuppressed
        End Function

    End Module
End Namespace