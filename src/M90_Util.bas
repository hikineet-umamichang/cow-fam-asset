Attribute VB_Name = "M90_Util"
Option Explicit

Public Sub ExportAllModules(ByVal exportPath As String)
    Dim vbComp As VBIDE.VBComponent
    Dim fso As Object
    Dim fileName As String

    If Len(Dir(exportPath, vbDirectory)) = 0 Then MkDir exportPath

    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule: fileName = exportPath & "\" & vbComp.Name & ".bas"
            Case vbext_ct_ClassModule: fileName = exportPath & "\" & vbComp.Name & ".cls"
            Case vbext_ct_MSForm: fileName = exportPath & "\" & vbComp.Name & ".frm"
            Case Else: fileName = ""
        End Select

        If fileName <> "" Then
            On Error Resume Next
            Kill fileName        ' ä˘ë∂ÇçÌèúÇµÇƒè„èëÇ´
            On Error GoTo 0
            vbComp.Export fileName
        End If
    Next vbComp
End Sub
