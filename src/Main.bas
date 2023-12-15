Attribute VB_Name = "Main"
Option Explicit

Dim swApp As Object
 
Sub Main()
    Dim currentDoc As ModelDoc2
    Dim selMgr As SelectionMgr
    Dim selCount As Integer
    Dim Comp As Component2
    Dim I As Integer
    Dim components As Dictionary
    Dim feat As Feature
     
    Set swApp = Application.SldWorks
    Set currentDoc = swApp.ActiveDoc
    If Not currentDoc Is Nothing Then
        If currentDoc.GetType = swDocASSEMBLY Then
            Set components = New Dictionary
            Set feat = currentDoc.FirstFeature
            Do Until feat Is Nothing
                Select Case feat.GetTypeName
                    Case "Reference"
                        Set Comp = feat.GetSpecificFeature2
                        If Not Comp.GetModelDoc2 Is Nothing Then 'supressed ignored
                            AddComponent components, Comp
                        End If
                    Case "FtrFolder"
                        ReorderComponents currentDoc, components
                        components.RemoveAll
                End Select
                Set feat = feat.GetNextFeature
            Loop
            ReorderComponents currentDoc, components
        End If
    End If
End Sub

Sub AddComponent(ByRef components As Dictionary, Comp As Component2)
    Dim compName As String
    Dim posMinus As Integer
    Dim key As String
    Dim Doc As ModelDoc2
    
    compName = Comp.Name2
    posMinus = InStrRev(compName, "-")
    key = BaseFilename(Comp.GetPathName) & "@" & Comp.ReferencedConfiguration & "@" _
        & Right(compName, Len(compName) - posMinus)
    components.Add key, Comp
End Sub

Function BaseFilename(pathname As String) As String
    Dim posSep As Integer
    
    posSep = InStrRev(pathname, "\")
    BaseFilename = Right(pathname, Len(pathname) - posSep)
End Function

Function SortAsmAndParts2(components As Dictionary) As Component2()
    Dim AsmArray() As String
    Dim PartArray() As String
    Dim Result() As Component2
    Dim AsmIndex As Integer
    Dim PartIndex As Integer
    Dim I As Integer
    Dim K As Variant
    Dim Comp As Component2
    Dim Doc As ModelDoc2
    
    ReDim AsmArray(components.Count - 1)
    AsmIndex = -1
    ReDim PartArray(components.Count - 1)
    PartIndex = -1
    For Each K In components.Keys
        Set Comp = components(K)
        Set Doc = Comp.GetModelDoc2
        If Doc.GetType = swDocASSEMBLY Then
            AsmIndex = AsmIndex + 1
            AsmArray(AsmIndex) = K
        Else
            PartIndex = PartIndex + 1
            PartArray(PartIndex) = K
        End If
    Next
    
    ReDim Result(components.Count - 1)
    If AsmIndex >= 0 Then
        ReDim Preserve AsmArray(AsmIndex)
        SortArray AsmArray
        For I = 0 To UBound(AsmArray)
            Set Result(I) = components(AsmArray(I))
        Next
    End If
    If PartIndex >= 0 Then
        ReDim Preserve PartArray(PartIndex)
        SortArray PartArray
        For I = 0 To UBound(PartArray)
            Set Result(AsmIndex + 1 + I) = components(PartArray(I))
        Next
    End If
    SortAsmAndParts2 = Result
End Function

Sub ReorderComponents(currentAsm As AssemblyDoc, components As Dictionary)
    Dim SortedComps As Variant
    
    If components.Count > 0 Then
        SortedComps = SortAsmAndParts2(components)
        currentAsm.ReorderComponents SortedComps, SortedComps(0), swReorderComponents_FirstInFolder
    End If
End Sub

Sub SortArray(ByRef arr As Variant)
    QuickSort arr, LBound(arr), UBound(arr)
End Sub

