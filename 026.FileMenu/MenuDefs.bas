Attribute VB_Name = "mMenuDefs"
Option Explicit

' Brought to you by Brad Martinez
'   http://members.aol.com/btmtz/vb
'   http://www.mvps.org/ccrp
'
' Code was written in and formatted for 8pt MS San Serif
'
' Note that "IShellFolder Extended Type Library v1.2" (ISHF_Ex.tlb)
' included with this project, must be present and correctly registered
' on your system, and referenced by this project, to allow use of the
' IShellFolder, IContextMenu and IMalloc interfaces.

' ====================================================

' C language BOOLEAN constants
Public Const CFalse = False
Public Const CTrue = 1

Public ICtxMenu2 As IShellFolderEx_TLB.IContextMenu2

Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
                      (pDest As Any, pSource As Any, ByVal dwLength As Long)

Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal Hwnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long

Public Const LB_ERR = -1
Public Const LB_SETSEL = &H185           ' multi-selection lbs only
Public Const LB_SETCURSEL = &H186   ' single selection lbs only
Public Const LB_GETCOUNT = &H18B
Public Const LB_SETCARETINDEX = &H19E  ' multi-selection lbs only

' Returns the listbox index if the specified point is over a list item,
' or - 1 otherwise. The ptX & ptY params want to be screen coords.
' Requires a tad more coding to make bAutoScroll functional but
' works nicely when dragging...
Declare Function LBItemFromPt Lib "comctl32.dll" _
                            (ByVal hLB As Long, _
                             ByVal ptX As Long, _
                             ByVal ptY As Long, _
                             ByVal bAutoScroll As Long) As Long

Public Type POINTAPI   ' pt
  x As Long
  y As Long
End Type

' Converts the specified window's client coordinates to screen coordinates
Declare Function ClientToScreen Lib "user32" _
                              (ByVal Hwnd As Long, _
                              lpPoint As POINTAPI) As Long

' ShowWindow commands
Public Enum SW_cmds
  SW_HIDE = 0
  SW_NORMAL = 1
  SW_SHOWNORMAL = 1
  SW_SHOWMINIMIZED = 2
  SW_MAXIMIZE = 3
  SW_SHOWMAXIMIZED = 3
  SW_SHOWNOACTIVATE = 4
  SW_SHOW = 5
  SW_MINIMIZE = 6
  SW_SHOWMINNOACTIVE = 7
  SW_SHOWNA = 8
  SW_RESTORE = 9
  SW_MAX = 10
  SW_SHOWDEFAULT = 10
End Enum

' ====================================================
' menu defs

Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function TrackPopupMenu Lib "user32" _
                              (ByVal hMenu As Long, _
                              ByVal wFlags As TPM_wFlags, _
                              ByVal x As Long, _
                              ByVal y As Long, _
                              ByVal nReserved As Long, _
                              ByVal Hwnd As Long, _
                              lprc As Any) As Long   ' lprc As RECT

Public Enum TPM_wFlags
  TPM_LEFTBUTTON = &H0
  TPM_RIGHTBUTTON = &H2
  TPM_LEFTALIGN = &H0
  TPM_CENTERALIGN = &H4
  TPM_RIGHTALIGN = &H8
  TPM_TOPALIGN = &H0
  TPM_VCENTERALIGN = &H10
  TPM_BOTTOMALIGN = &H20

  TPM_HORIZONTAL = &H0         ' Horz alignment matters more
  TPM_VERTICAL = &H40            ' Vert alignment matters more
  TPM_NONOTIFY = &H80           ' Don't send any notification msgs
  TPM_RETURNCMD = &H100
End Enum

Public Type MENUITEMINFO
  cbSize As Long
  fMask As MII_Mask
  fType As MF_Type              ' MIIM_TYPE
  fState As MF_State             ' MIIM_STATE
  wID As Long                       ' MIIM_ID
  hSubMenu As Long            ' MIIM_SUBMENU
  hbmpChecked As Long      ' MIIM_CHECKMARKS
  hbmpUnchecked As Long  ' MIIM_CHECKMARKS
  dwItemData As Long          ' MIIM_DATA
  dwTypeData As String        ' MIIM_TYPE
  cch As Long                       ' MIIM_TYPE
End Type

Public Enum MII_Mask
  MIIM_STATE = &H1
  MIIM_ID = &H2
  MIIM_SUBMENU = &H4
  MIIM_CHECKMARKS = &H8
  MIIM_TYPE = &H10
  MIIM_DATA = &H20
End Enum

' win40  -- A lot of MF_* flags have been renamed as MFT_* and MFS_* flags
Public Enum MenuFlags
  MF_INSERT = &H0
  MF_ENABLED = &H0
  MF_UNCHECKED = &H0
  MF_BYCOMMAND = &H0
  MF_STRING = &H0
  MF_UNHILITE = &H0
  MF_GRAYED = &H1
  MF_DISABLED = &H2
  MF_BITMAP = &H4
  MF_CHECKED = &H8
  MF_POPUP = &H10
  MF_MENUBARBREAK = &H20
  MF_MENUBREAK = &H40
  MF_HILITE = &H80
  MF_CHANGE = &H80
  MF_END = &H80                    ' Obsolete -- only used by old RES files
  MF_APPEND = &H100
  MF_OWNERDRAW = &H100
  MF_DELETE = &H200
  MF_USECHECKBITMAPS = &H200
  MF_BYPOSITION = &H400
  MF_SEPARATOR = &H800
  MF_REMOVE = &H1000
  MF_DEFAULT = &H1000
  MF_SYSMENU = &H2000
  MF_HELP = &H4000
  MF_RIGHTJUSTIFY = &H4000
  MF_MOUSESELECT = &H8000&
End Enum

Public Enum MF_Type
  MFT_STRING = MF_STRING
  MFT_BITMAP = MF_BITMAP
  MFT_MENUBARBREAK = MF_MENUBARBREAK
  MFT_MENUBREAK = MF_MENUBREAK
  MFT_OWNERDRAW = MF_OWNERDRAW
  MFT_RADIOCHECK = &H200
  MFT_SEPARATOR = MF_SEPARATOR
  MFT_RIGHTORDER = &H2000
  MFT_RIGHTJUSTIFY = MF_RIGHTJUSTIFY
End Enum

Public Enum MF_State
  MFS_GRAYED = &H3
  MFS_DISABLED = MFS_GRAYED
  MFS_CHECKED = MF_CHECKED
  MFS_HILITE = MF_HILITE
  MFS_ENABLED = MF_ENABLED
  MFS_UNCHECKED = MF_UNCHECKED
  MFS_UNHILITE = MF_UNHILITE
  MFS_DEFAULT = MF_DEFAULT
End Enum

Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" _
                              (ByVal hMenu As Long, _
                              ByVal uItem As Long, _
                              ByVal fByPosition As Boolean, _
                              lpmii As MENUITEMINFO) As Boolean

Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" _
                              (ByVal hMenu As Long, _
                              ByVal uItem As Long, _
                              ByVal fByPosition As Boolean, _
                              lpmii As MENUITEMINFO) As Boolean
'

' Displays the specified items' shell context menu.
'
'    hwndOwner  - window handle that owns context menu and any err msgboxes
'    isfParent       - pointer to the items' parent shell folder
'    cPidls            - count of pidls at, and after, pidlRel
'    pidlRel          - the first item's pidl, relative to isfParent
'    pt                  - location of the context menu, in screen coords
'    fPrompt         - flag specifying whether to prompt before executing any selected
'                           context menu command
'
' Returns True if a context menu command was selected, False otherwise.

Public Function ShowShellContextMenu(hwndOwner As Long, _
                                                                isfParent As IShellFolderEx_TLB.IShellFolder, _
                                                                cPidls As Integer, _
                                                                pidlRel As Long, _
                                                                pt As POINTAPI, _
                                                                fPrompt As Boolean) As String
  Dim IID_IContextMenu As IShellFolderEx_TLB.GUID
  Dim IID_IContextMenu2 As IShellFolderEx_TLB.GUID
  Dim icm As IShellFolderEx_TLB.IContextMenu
  Dim hr As Long   ' HRESULT
  Dim hMenu As Long
  Dim idCmd As Long
  Dim cmi As IShellFolderEx_TLB.CMINVOKECOMMANDINFO
  
  ' Fill the IContextMenu interface ID, {000214E4-000-000-C000-000000046}
  With IID_IContextMenu
    .Data1 = &H214E4
    .Data4(0) = &HC0
    .Data4(7) = &H46
  End With
    
  ' Get a refernce to the item's IContextMenu interface.
  hr = isfParent.GetUIObjectOf(hwndOwner, cPidls, pidlRel, IID_IContextMenu, 0, icm)
  If hr >= NOERROR Then
    
    ' Fill the IContextMenu2 interface ID, {000214F4-000-000-C000-000000046}
    ' and get the folder's IContextMenu2. Is needed so the "Send To" and "Open
    ' With" submenus get filled from the HandleMenuMsg call in WndProc.
    With IID_IContextMenu2
      .Data1 = &H214F4
      .Data4(0) = &HC0
      .Data4(7) = &H46
    End With
    Call icm.QueryInterface(IID_IContextMenu2, ICtxMenu2)
    
    ' Create a new popup menu...
    hMenu = CreatePopupMenu()
    If hMenu Then

      ' Add the item's shell commands to the popup menu.
      If (ICtxMenu2 Is Nothing) = False Then
        hr = ICtxMenu2.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_EXPLORE)
      Else
        hr = icm.QueryContextMenu(hMenu, 0, 1, &H7FFF, CMF_EXPLORE)
      End If
      If hr >= NOERROR Then
        
        ' Show the item's context menu
        idCmd = TrackPopupMenu(hMenu, _
                                                    TPM_LEFTALIGN Or _
                                                    TPM_RETURNCMD Or _
                                                    TPM_RIGHTBUTTON, _
                                                    pt.x, pt.y, 0, hwndOwner, 0)
        
        ' If a menu command is selected...
        If idCmd Then
        
          ' If still executing the command...
          If idCmd Then
            
            ShowShellContextMenu = GetMenuCmdStr(hMenu, (idCmd))
            ' Fill the struct with the selected command's information.
            With cmi
              .cbSize = Len(cmi)
              .Hwnd = hwndOwner
              .lpVerb = idCmd - 1 ' MAKEINTRESOURCE(idCmd-1);
              .nShow = SW_SHOWNORMAL
            End With
  
            ' Invoke the shell's context menu command. The call itself does
            ' not err if the pidlRel item is invalid, but depending on the selected
            ' command, Explorer *may* raise an err. We don't need the return
            ' val, which should always be NOERROR anyway...
            If (ICtxMenu2 Is Nothing) = False Then
              Call ICtxMenu2.InvokeCommand(cmi)
            Else
              Call icm.InvokeCommand(cmi)
            End If
          
          End If   ' idCmd
        End If   ' idCmd
      End If   ' hr >= NOERROR (QueryContextMenu)

      Call DestroyMenu(hMenu)
    
    End If   ' hMenu
  End If   ' hr >= NOERROR (GetUIObjectOf)

  ' Release the folder's IContextMenu2 from the global variable.
  Set ICtxMenu2 = Nothing
  
  ' Return True if a menu command was selected
  ' (letting us know to react accordingly...)

End Function

' Returns the string of the specified menu command ID in the specified menu.

Public Function GetMenuCmdStr(hMenu As Long, idCmd As Integer) As String
  Dim mii As MENUITEMINFO
  
  ' Initialize the struct..
  With mii
    .cbSize = Len(mii)
    .fMask = MIIM_TYPE
    .fType = MFT_STRING
    .dwTypeData = String$(256, 0)
    .cch = 256
  End With
  
  ' Returns TRUE on success
  If GetMenuItemInfo(hMenu, idCmd, False, mii) Then
    GetMenuCmdStr = Left$(mii.dwTypeData, mii.cch)
  End If

End Function
