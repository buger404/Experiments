Attribute VB_Name = "mShellDefs"
Option Explicit

' Brought to you by Brad Martinez
'   http://members.aol.com/btmtz/vb
'   http://www.mvps.org/ccrp
'
' Code was written in and formatted for 8pt MS San Serif
'
' Note that "IShellFolder Extended Type Library v1.1" (ISHF_Ex.tlb)
' included with this project, must be present and correctly registered
' on your system, and referenced by this project, to allow use of the
' IShellFolder, IContextMenu and IMalloc interfaces.

' ====================================================

' Defined as an HRESULT that corresponds to S_OK.
Public Const NOERROR = 0

' Retrieves the IShellFolder interface for the desktop folder.
' Returns NOERROR if successful or an OLE-defined error result otherwise.
Declare Function SHGetDesktopFolder Lib "shell32" (ppshf As IShellFolder) As Long

' Retrieves a pointer to the shell's IMalloc interface.
' Returns NOERROR if successful or or E_FAIL otherwise.
Declare Function SHGetMalloc Lib "shell32" (ppMalloc As IMalloc) As Long

' GetItemID item ID retrieval constants
Public Const GIID_FIRST = 1
Public Const GIID_LAST = -1
'
' ====================================================
' item ID (pidl) structs, just for reference
'
'' item identifier (relative pidl), allocated by the shell
'Type SHITEMID
'  cb As Integer        ' size of struct, including cb itself
'  abID(0) As Byte    ' variable length item identifier
'End Type
'
'' fully qualified pidl
'Type ITEMIDLIST
'  mkid As SHITEMID  ' list of item identifers, packed into SHITEMID.abID
'End Type
'

' Returns a reference to the IMalloc interface.

Public Function MemAllocator() As IShellFolderEx_TLB.IMalloc
  Static im As IShellFolderEx_TLB.IMalloc
  ' SHGetMalloc should just get called once as the 'im'
  ' variable stays in scope while the project is running...
  If im Is Nothing Then Call SHGetMalloc(im)
  Set MemAllocator = im
End Function

' ====== Begin pidl procs ===============================

' Determines if the specified pidl is the desktop folder's pidl.
' Returns True if the pidl is the desktop's pidl, returns False otherwise.

' The desktop pidl is only a single item ID whose value is 0 (the 2 byte
' zero-terminator, i.e. SHITEMID.abID is empty). Direct descendents of
' the desktop (My Computer, Network Neighborhood) are absolute pidls
' (relative to the desktop) also with a single item ID, but contain values
' (SHITEMID.abID > 0). Drive folders have 2 item IDs, children of drive
' folders have 3 item IDs, etc. All other single item ID pidls are relative to
' the shell folder in which they reside (just like a relative path).

Public Function IsDesktopPIDL(pidl As Long) As Boolean
  ' The GetItemIDSize() call will also return 0 if pidl = 0
  If pidl Then IsDesktopPIDL = (GetItemIDSize(pidl) = 0)
End Function

' Returns the size in bytes of the first item ID in a pidl.
' Returns 0 if the pidl is the desktop's pidl or is the last
' item ID in the pidl (the zero terminator), or is invalid.

Public Function GetItemIDSize(ByVal pidl As Long) As Integer
  ' If we try to access memory at address 0 (NULL), then it's bye-bye...
  If pidl Then MoveMemory GetItemIDSize, ByVal pidl, 2
End Function

' Returns the count of item IDs in a pidl.

Public Function GetItemIDCount(ByVal pidl As Long) As Integer
  Dim nItems As Integer
  ' If the size of an item ID is 0, then it's the zero
  ' value terminating item ID at the end of the pidl.
  Do While GetItemIDSize(pidl)
    pidl = GetNextItemID(pidl)
    nItems = nItems + 1
  Loop
  GetItemIDCount = nItems
End Function

' Returns a pointer to the next item ID in a pidl.
' Returns 0 if the next item ID is the pidl's zero value terminating 2 bytes.

Public Function GetNextItemID(ByVal pidl As Long) As Long
  Dim cb As Integer   ' SHITEMID.cb, 2 bytes
  cb = GetItemIDSize(pidl)
  ' Make sure it's not the zero value terminator.
  If cb Then GetNextItemID = pidl + cb
End Function

' If successful, returns the size in bytes of the memory occcupied by a pidl,
' including it's 2 byte zero terminator. Returns 0 otherwise.

Public Function GetPIDLSize(ByVal pidl As Long) As Integer
  Dim cb As Integer
  ' Error handle in case we get a bad pidl and overflow cb.
  ' (most item IDs are roughly 20 bytes in size, and since an item ID represents
  ' a folder, a pidl can never exceed 260 folders, or 5200 bytes).
  On Error GoTo Out
  
  If pidl Then
    Do While pidl
      cb = cb + GetItemIDSize(pidl)
      pidl = GetNextItemID(pidl)
    Loop
    ' Add 2 bytes for the zero terminating item ID
    GetPIDLSize = cb + 2
  End If
  
Out:
End Function

' Copies and returns the specified item ID from a complex pidl
'   pidl -    pointer to an item ID list from which to copy
'   nItem - 1-based position in the pidl of the item ID to copy

' If successful, returns a new item ID (single-element pidl)
' from the specified element positon. Returns 0 on failure.
' If nItem exceeds the number of item IDs in the pidl,
' the last item ID is returned.
' (calling proc is responsible for freeing the new pidl)

Public Function GetItemID(ByVal pidl As Long, ByVal nItem As Integer) As Long
  Dim nCount As Integer
  Dim i As Integer
  Dim cb As Integer
  Dim pidlNew As Long
  
  nCount = GetItemIDCount(pidl)
  If (nItem > nCount) Or (nItem = GIID_LAST) Then nItem = nCount
  
  ' GetNextItemID returns the 2nd item ID
  For i = 1 To nItem - 1: pidl = GetNextItemID(pidl): Next
    
  ' Get the size of the specified item identifier.
  ' If cb = 0 (the zero terminator), the we'll return a desktop pidl, proceed
  cb = GetItemIDSize(pidl)
  
  ' Allocate a new item identifier list.
  pidlNew = MemAllocator.Alloc(cb + 2)
  If pidlNew Then
    
    ' Copy the specified item identifier.
    ' and append the zero terminator.
    MoveMemory ByVal pidlNew, ByVal pidl, cb
    MoveMemory ByVal pidlNew + cb, 0, 2
    
    GetItemID = pidlNew
  End If
  
End Function

' Returns an absolute pidl (relative to the desktop) from a valid file system
' path only (i.e. not from a display name).

'   hwndOwner - handle of window that will own any displayed msg boxes
'   sPath           - fully qualified path whose pidl is to be returned

' If successful, the path's pidl is returned, otherwise 0 is returned.
' (calling proc is responsible for freeing the pidl)

Public Function GetPIDLFromPath(hwndOwner As Long, _
                                                      sPath As String) As Long
  Dim isfDesktop As IShellFolderEx_TLB.IShellFolder
  Dim pchEaten As Long
  Dim pidl As Long

  If SHGetDesktopFolder(isfDesktop) = NOERROR Then
    If isfDesktop.ParseDisplayName(hwndOwner, 0, _
                                                        StrConv(sPath, vbUnicode), _
                                                        pchEaten, _
                                                        pidl, 0) = NOERROR Then
      GetPIDLFromPath = pidl
    End If
  End If
End Function
'
' ====== End pidl procs ===============================
'

' Returns a reference to the parent IShellFolder of the last
' item ID in the specified fully qualified pidl.

' If pidlFQ is zero, or a relative (single item) pidl, then the
' desktop's IShellFolder is returned.
' If an unexpected error occurs, the object value Nothing is returned.

Public Function GetParentIShellFolder(ByVal pidlFQ As Long) As IShellFolder
  Dim nCount As Integer
  Dim i As Integer
  Dim isf As IShellFolderEx_TLB.IShellFolder
  Dim pidlRel As Long
  Dim IID_IShellFolder As IShellFolderEx_TLB.GUID
  On Error GoTo Out

  nCount = GetItemIDCount(pidlFQ)
  ' If pidlFQ is 0 and is not the desktop's pidl...
  If (nCount = 0) And (IsDesktopPIDL(pidlFQ) = False) Then Error 1
  
  ' Get the desktop's IShellfolder first.
  If SHGetDesktopFolder(isf) = NOERROR Then
    
    ' Fill the IShellFolder interface ID, {000214E6-000-000-C000-000000046}
    With IID_IShellFolder
      .Data1 = &H214E6
      .Data4(0) = &HC0
      .Data4(7) = &H46
    End With
    
    ' Walk through the pidl and bind all the way to it's *2nd to last* item ID.
    For i = 1 To nCount - 1
      
      ' Get the next item ID in the pidl (child of the current IShellFolder)
      pidlRel = GetItemID(pidlFQ, i)
      
      ' Bind to the item current ID's folder and get it's IShellFolder
      If isf.BindToObject(pidlRel, 0, IID_IShellFolder, isf) <> NOERROR Then Error 1
      
      ' Free the current item ID and zero it
      MemAllocator.Free ByVal pidlRel
      pidlRel = 0
    
    Next
  
  End If   ' SHGetDesktopFolder(isf) = NOERROR
  
Out:
  If pidlRel Then MemAllocator.Free ByVal pidlRel
  
  ' Return a reference to the IShellFolder
  Set GetParentIShellFolder = isf
  
End Function
