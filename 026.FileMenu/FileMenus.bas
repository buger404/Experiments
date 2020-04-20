Attribute VB_Name = "FileMenus"
'==============================================
'   ATTENTION
'==============================================
'   YOU NEED:
'   MenuDefs.bas , ShellDefs.bas , ISHF_Ex.tlb
'   Return Value:
'   Clicked menu ("" = cancel)
'==============================================
Function ShowFileMenu(ByVal Path As String, ByVal hWnd As Long) As String  ' (FilePath / Window Hwnd)

    Dim p As POINT, pt As POINTAPI, i As Integer
    Dim fq As Long, isfParent As IShellFolder, rel As Long
    
    GetCursorPos p: pt.x = p.x: pt.y = p.y ' Get cursor position
    
    fq = GetPIDLFromPath(hWnd, Path) 'Get FQ
    
    If fq = 0 Then Exit Function
        
    Set isfParent = GetParentIShellFolder(fq) 'I don't know what's this
    If isfParent Is Nothing Then Exit Function
    
    rel = GetItemID(fq, GIID_LAST) 'Get Rel
            
    If rel Then
        ShowFileMenu = ShowShellContextMenu(hWnd, isfParent, 1, rel, pt, True)  'Show
    End If
        
    'Dustbin ©³(£Þ0£Þ)©¿
    MemAllocator.Free ByVal rel: MemAllocator.Free ByVal fq
    
End Function
