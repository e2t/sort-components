Attribute VB_Name = "Main"
Option Explicit

Dim swApp As Object
 
Sub Main()
    Dim currentDoc As ModelDoc2
    Dim selMgr As SelectionMgr
    Dim selCount As Integer
    Dim comp As Component2
    Dim i As Integer
    Dim components As Dictionary
    Dim feat As Feature
     
    Set swApp = Application.SldWorks
    Set currentDoc = swApp.ActiveDoc
    If Not currentDoc Is Nothing Then
        If currentDoc.GetType = swDocASSEMBLY Then
            Set components = New Dictionary
            'Set selMgr = currentDoc.SelectionManager
            'selCount = selMgr.GetSelectedObjectCount2(-1)
            'If selCount > 0 Then
            '    For i = 1 To selCount
            '        Set comp = selMgr.GetSelectedObjectsComponent4(i, -1)
            '        If Not comp Is Nothing Then
            '            AddComponent components, comp
            '        End If
            '    Next
            '    ReorderComponents currentDoc, components
            'Else
                Set feat = currentDoc.FirstFeature
                Do Until feat Is Nothing
                    Select Case feat.GetTypeName
                        Case "Reference"
                            Set comp = feat.GetSpecificFeature2
                            AddComponent components, comp
                        Case "FtrFolder"
                            ReorderComponents currentDoc, components
                            components.RemoveAll
                    End Select
                    Set feat = feat.GetNextFeature
                Loop
                ReorderComponents currentDoc, components
            'End If
        End If
    End If
End Sub

Sub AddComponent(ByRef components As Dictionary, comp As Component2)
    Dim compName As String
    Dim posMinus As Integer
    Dim key As String
    Dim doc As ModelDoc2
    
    compName = comp.Name2
    posMinus = InStrRev(compName, "-")
    key = BaseFilename(comp.GetPathName) & "@" & comp.ReferencedConfiguration & "@" & Right(compName, Len(compName) - posMinus)
    components.Add key, comp
End Sub

Function BaseFilename(pathname As String) As String
    Dim posSep As Integer
    
    posSep = InStrRev(pathname, "\")
    BaseFilename = Right(pathname, Len(pathname) - posSep)
End Function

Function SortAsmAndParts(components As Dictionary) As String()
    Dim res() As String
    Dim assemblies() As String
    Dim parts() As String
    Dim asmCount As Integer
    Dim partCount As Integer
    Dim i As Variant
    Dim comp As Component2
    Dim doc As ModelDoc2
    Dim j As Integer
    
    ReDim res(components.Count)
    ReDim assemblies(components.Count)
    ReDim parts(components.Count)
    asmCount = -1
    partCount = -1
    For Each i In components.Keys
        Set comp = components(i)
        Set doc = comp.GetModelDoc2
        If doc.GetType = swDocASSEMBLY Then
            asmCount = asmCount + 1
            assemblies(asmCount) = i
        Else
            partCount = partCount + 1
            parts(partCount) = i
        End If
    Next
    j = -1
    If asmCount >= 0 Then
        ReDim Preserve assemblies(asmCount)
        assemblies = BubbleSort(assemblies)
        For asmCount = LBound(assemblies) To UBound(assemblies)
            j = j + 1
            res(j) = assemblies(asmCount)
        Next
    End If
    If partCount >= 0 Then
        ReDim Preserve parts(partCount)
        parts = BubbleSort(parts)
        For partCount = LBound(parts) To UBound(parts)
            j = j + 1
            res(j) = parts(partCount)
        Next
    End If
    SortAsmAndParts = res
End Function

Sub ReorderComponents(currentAsm As AssemblyDoc, components As Dictionary)
    Dim sortedKeys As Variant
    Dim i As Integer
    
    sortedKeys = SortAsmAndParts(components)
    For i = LBound(sortedKeys) + 1 To UBound(sortedKeys)
        currentAsm.ReorderComponents components(sortedKeys(i)), components(sortedKeys(i - 1)), swReorderComponents_After
    Next
End Sub

Function BubbleSort(ByVal arr As Variant) As Variant
    Dim i As Integer
    Dim j As Integer
    Dim tmp As Variant
    Dim needSort As Boolean
    Dim lowIndex As Integer
    Dim upIndex As Integer
    
    lowIndex = LBound(arr)
    upIndex = UBound(arr) - 1
    
    needSort = True
    i = lowIndex
    While i <= upIndex And needSort
        needSort = False
        For j = lowIndex To upIndex - i
            If arr(j) > arr(j + 1) Then
                tmp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = tmp
                needSort = True
            End If
        Next
        i = i + 1
    Wend
    BubbleSort = arr
End Function
