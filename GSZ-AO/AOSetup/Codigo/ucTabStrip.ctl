VERSION 5.00
Begin VB.UserControl ucTabStrip 
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   PropertyPages   =   "ucTabStrip.ctx":0000
   ScaleHeight     =   73
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   130
   ToolboxBitmap   =   "ucTabStrip.ctx":0012
End
Attribute VB_Name = "ucTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ------------------------------------------------------------------------
' Author:        Raul E. Arellano
'                Started by Leandro Ascierto (www.leandroascierto.com.ar)
' Dependencies:  ppgTabStrip.pag -> Purely recommended if you put controls in Design Mode (to add/change/set tabs)
'                ucTabStrip.ctx -> The Toolbar Icon :D
' History:       July 11, 2011...................First Cut
'                December 10, 2011...............Added suppourt for control array
'                                                Fixed crash when InnerControls > 30
' ------------------------------------------------------------------------
'    DO NOT INSERT CONTROLS WITHOUT hWnd PROPERTY!!! (Handle it outside the control by code)
' ------------------------------------------------------------------------
' Thanks To
'     Self-Subclassing UserControl template (IDE safe).
'
'     From original post by LaVolpe (Worked from Paul Caton code)
'
'     Self-subclassing Controls/Forms - NO dependencies
'     http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=68737&lngWId=1
'----------------------------------------------------------------------------------------
'     Traping TabStop + navigation keys. (Reformed by Raul338 to not depend from a module)
'
'     From original post by Vlad Vissoultchev:
'
'     How to capture Tab/Enter/Esc on your custom UserControl
'     http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=41506&lngWId=1
' ------------------------------------------------------------------------
Option Explicit
Option Base 0

' Events
Public Event Click()
Public Event DblClick()
Public Event ChangingTab(Cancel As Boolean)
Public Event TabClick(ByVal lTab As Long)
Public Event TabRightClick(ByVal lTab As Long)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Public Event MouseEnter()
Public Event MouseLeave()

' Private Variables
' XP Theme and Tab Handle
Private hTabs               As Long
Private hMod                As Long
' Self Subclassing
Private hOldWndProc         As Long
Private hOldTabWndProc      As Long
Private hCallBackHandle     As Long
' Mouse in control flag, coords for mouse events
Private m_bInCtrl           As Boolean
Private m_snxL              As Single
Private m_snyL              As Single
' Font
Private m_hFont             As Long
Private WithEvents m_oFont  As StdFont
Attribute m_oFont.VB_VarHelpID = -1

' Tab designer
Private Const m_Splitter    As String = "###Ђ"
Private Type udtControl
    TabIndex As Long
    Name     As String
End Type
Private m_controlsTab()     As udtControl

' === Windows General ====================================================
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function InitCommonControls Lib "COMCTL32" () As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function PtInRect Lib "user32" (lprc As RECT, lpPoint As POINT) As Boolean
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINT) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal Visible As Long) As Boolean
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Boolean
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const GWL_STYLE = (-16)

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOZORDER      As Long = &H4

Private Declare Function CreateFontIndirect Lib "GDI32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal HDC As Long, ByVal nIndex As Long) As Long
Private Type LOGFONT
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQuality        As Byte
    lfPitchAndFamily As Byte
    lfFaceName(32)   As Byte
End Type

Private Const LOGPIXELSY             As Long = 90
Private Const FW_NORMAL              As Long = 400
Private Const FW_BOLD                As Long = 700
Private Const FF_DONTCARE            As Long = 0
Private Const DEFAULT_QUALITY        As Long = 0
Private Const DEFAULT_PITCH          As Long = 0
Private Const DEFAULT_CHARSET        As Long = 1
Private Const NONANTIALIASED_QUALITY As Long = 3

Private Enum TRACKMOUSEEVENT_FLAGS
    [TME_HOVER] = &H1&
    [TME_LEAVE] = &H2&
    [TME_QUERY] = &H40000000
    [TME_CANCEL] = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize      As Long
    dwFlags     As TRACKMOUSEEVENT_FLAGS
    hwndTrack   As Long
    dwHoverTime As Long
End Type

' Mensajes
Private Const WM_DESTROY            As Long = &H2
Private Const WM_SETFOCUS           As Long = &H7
Private Const WM_SETFONT            As Long = &H30
Private Const WM_NOTIFY             As Long = &H4E
Private Const WM_MOUSEACTIVATE      As Long = &H21
Private Const WM_KEYDOWN            As Long = &H100
Private Const WM_KEYUP              As Long = &H101
Private Const WM_CHAR               As Long = &H102
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_LBUTTONUP          As Long = &H202
Private Const WM_LBUTTONDOWN        As Long = &H201
Private Const WM_RBUTTONDOWN        As Long = &H204
Private Const WM_RBUTTONUP          As Long = &H205
Private Const WM_MBUTTONDOWN        As Long = &H207
Private Const WM_MBUTTONUP          As Long = &H208
Private Const WM_MOUSELEAVE         As Long = &H2A3
' Notifications
Private Const NM_FIRST              As Long = 0
Private Const NM_CLICK              As Long = (NM_FIRST - 2)
Private Const NM_DBLCLK             As Long = (NM_FIRST - 3)
Private Const NM_RETURN             As Long = (NM_FIRST - 4)
Private Const NM_RCLICK             As Long = (NM_FIRST - 5)
Private Const NM_RDBLCLK            As Long = (NM_FIRST - 6)
' Styles
Private Const WS_CHILD              As Long = &H40000000
Private Const WS_CLIPCHILDREN       As Long = &H2000000
Private Const WS_CLIPSIBLINGS       As Long = &H4000000
Private Const WS_OVERLAPPED         As Long = &H0&
Private Const WS_VISIBLE            As Long = &H10000000
Private Const WS_TABS As Long = (WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_OVERLAPPED Or WS_VISIBLE Or WS_CHILD)
' Ex-Styles
Private Const WS_EX_LEFT            As Long = &H0&
Private Const WS_EX_LTRREADING      As Long = &H0&
Private Const WS_EX_RIGHTSCROLLBAR  As Long = &H0&
Private Const WS_EX_TABS = (WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR)

Private Type NMHDR
    hwndFrom As Long
    idfrom   As Long
    code     As Long
End Type
Private Type POINT
    X       As Long
    y       As Long
End Type
Private Type RECT
    Left                    As Long
    Top                     As Long
    Right                   As Long
    Bottom                  As Long
End Type

' ==== TabStrip ==========================================================
' Mensajes
Private Const TCM_FIRST             As Long = &H1300
Private Const TCM_GETIMAGELIST      As Long = (TCM_FIRST + 2)
Private Const TCM_SETIMAGELIST      As Long = (TCM_FIRST + 3)
Private Const TCM_GETITEMCOUNT      As Long = (TCM_FIRST + 4)
Private Const TCM_INSERTITEM        As Long = (TCM_FIRST + 7)
Private Const TCM_DELETEITEM        As Long = (TCM_FIRST + 8)
Private Const TCM_DELETEALLITEMS    As Long = (TCM_FIRST + 9)
Private Const TCM_GETITEMRECT       As Long = (TCM_FIRST + 10)
Private Const TCM_GETCURSEL         As Long = (TCM_FIRST + 11)
Private Const TCM_SETCURSEL         As Long = (TCM_FIRST + 12)
Private Const TCM_HITTEST           As Long = (TCM_FIRST + 13)
Private Const TCM_SETITEMEXTRA      As Long = (TCM_FIRST + 14)
Private Const TCM_ADJUSTRECT        As Long = (TCM_FIRST + 40)
Private Const TCM_SETITEMSIZE       As Long = (TCM_FIRST + 41)
Private Const TCM_REMOVEIMAGE       As Long = (TCM_FIRST + 42)
Private Const TCM_SETPADDING        As Long = (TCM_FIRST + 43)
Private Const TCM_GETROWCOUNT       As Long = (TCM_FIRST + 44)
Private Const TCM_GETTOOLTIPS       As Long = (TCM_FIRST + 45)
Private Const TCM_SETTOOLTIPS       As Long = (TCM_FIRST + 46)
Private Const TCM_GETCURFOCUS       As Long = (TCM_FIRST + 47)
Private Const TCM_SETCURFOCUS       As Long = (TCM_FIRST + 48)
Private Const TCM_SETMINTABWIDTH    As Long = (TCM_FIRST + 49)
Private Const TCM_DESELECTALL       As Long = (TCM_FIRST + 50)
Private Const TCM_HIGHLIGHTITEM     As Long = (TCM_FIRST + 51)
Private Const TCM_SETEXTENDEDSTYLE  As Long = (TCM_FIRST + 52)
Private Const TCM_GETEXTENDEDSTYLE  As Long = (TCM_FIRST + 53)
Private Const TCM_GETITEMW          As Long = (TCM_FIRST + 60)
Private Const TCM_SETITEMW          As Long = (TCM_FIRST + 61)
Private Const TCM_INSERTITEMW       As Long = (TCM_FIRST + 62)
' Styles
Private Const TCS_SINGLELINE        As Long = &H0
Private Const TCS_RIGHTJUSTIFY      As Long = &H0
Private Const TCS_TABS              As Long = &H0
Private Const TCS_SCROLLOPPOSITE    As Long = &H1
Private Const TCS_RIGHT             As Long = &H2
Private Const TCS_BOTTOM            As Long = &H2
Private Const TCS_MULTISELECT       As Long = &H4
Private Const TCS_FLATBUTTONS       As Long = &H8
Private Const TCS_FORCEICONLEFT     As Long = &H10
Private Const TCS_FORCELABELLEFT    As Long = &H20
Private Const TCS_HOTTRACK          As Long = &H40
Private Const TCS_VERTICAL          As Long = &H80
Private Const TCS_BUTTONS           As Long = &H100
Private Const TCS_MULTILINE         As Long = &H200
Private Const TCS_FIXEDWIDTH        As Long = &H400
Private Const TCS_RAGGEDRIGHT       As Long = &H800
Private Const TCS_FOCUSNEVER        As Long = &H8000
Private Const TCS_FOCUSONBUTTONDOWN As Long = &H1000
Private Const TCS_OWNERDRAWFIXED    As Long = &H2000
Private Const TCS_TOOLTIPS          As Long = &H4000
' Ex-Styles
Private Const TCS_EX_FLATSEPARATORS As Long = &H1
Private Const TCS_EX_REGISTERDROP   As Long = &H2
' HitTest
Private Const TCHT_ONITEMICON       As Long = &H2
Private Const TCHT_ONITEMLABEL      As Long = &H4
Private Const TCHT_NOWHERE          As Long = &H1
Private Const TCHT_ONITEM           As Long = (TCHT_ONITEMICON Or TCHT_ONITEMLABEL)
' Item Flags
Private Const TCIF_IMAGE            As Long = &H2
Private Const TCIF_PARAM            As Long = &H8
Private Const TCIF_RTLREADING       As Long = &H4
Private Const TCIF_STATE            As Long = &H10
Private Const TCIF_TEXT             As Long = &H1
' Item States
Private Const TCIS_BUTTONPRESSED    As Long = &H1
Private Const TCIS_HIGHLIGHTED      As Long = &H2
' Notifications
Private Const TCN_FIRST             As Long = -550
Private Const TCN_SELCHANGE         As Long = (TCN_FIRST - 1)
Private Const TCN_SELCHANGING       As Long = (TCN_FIRST - 2)
Private Const TCN_FOCUSCHANGE       As Long = (TCN_FIRST - 4)

Private Const WC_TABCONTROL         As String = "SysTabControl32"

Private Type TCHITTESTINFO
    pt          As POINT
    flags       As Long
End Type

Private Type TCITEM
    mask        As Long
    dwState     As Long
    dwStateMask As Long
    pszText     As Long
    cchTextMax  As Long
    iImage      As Long
    lParam      As Long
End Type

'========================================================================================
' mIOLEInPlaceActiveObject Implementation
' Author:      Mike Gainer, Matt Curland and Bill Storage
'
' Requires:    OleGuids.tlb (in IDE only)
'
' Description:
' Allows you to replace the standard IOLEInPlaceActiveObject interface for a
' UserControl with a customisable one.  This allows you to take control
' of focus in VB controls.
'
' The code could be adapted to replace other UserControl OLE interfaces.
'
' ---------------------------------------------------------------------------------------
' Visit vbAccelerator, advanced, free source for VB programmers
' http://vbaccelerator.com
'========================================================================================
Private Type IPAOHookStruct
    lpVTable    As Long                    'VTable pointer
    IPAOReal    As Long 'IOleInPlaceActiveObject 'Un-AddRefed pointer for forwarding calls
    ThisPointer As Long
End Type
Private m_uIPAO         As IPAOHookStruct
Private Declare Function IsEqualGUID Lib "ole32" (iid1 As uuid, iid2 As uuid) As Long

Private Type OLEINPLACEFRAMEINFO
    cb              As Long
    fMDIApp         As Boolean
    hwndFrame       As Long
    haccel          As Long
    cAccelEntries   As Long
End Type

Private Type msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINT
End Type 'MSG

Private Const S_FALSE               As Long = 1
Private Const S_OK                  As Long = 0

Private IID_IOleInPlaceActiveObject As uuid
Private m_IPAOVTable(9)             As Long

'*************************************************************************************************
' ==== Used by CallInterface Function =====================================================

Private Type uuid
  Data1         As Long
  Data2         As Integer
  Data3         As Integer
  Data4(0 To 7) As Byte
End Type

Private Enum IUnknown_Exports
    [QueryInterface] = 0
    [AddRef] = 1
    [Release] = 2
End Enum

Private Enum IPAO_Exports
    [GetWindow] = 3
    [ContextSensitiveHelp] = 4
    [TranslateAccelerator] = 5
    [OnFrameWindowActivate] = 6
    [OnDocWindowActivate] = 7
    [ResizeBorder] = 8
    [EnableModeless] = 9
End Enum

Private Declare Function PutMem2 Lib "msvbvm60" (ByVal pWORDDst As Long, ByVal NewValue As Long) As Long
Private Declare Function PutMem4 Lib "msvbvm60" (ByVal pDWORDDst As Long, ByVal NewValue As Long) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByVal pDWORDSrc As Long, ByVal pDWORDDst As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, lpiid As uuid) As Long

Private Const IID_IOleInPlaceActive     As String = "{00000117-0000-0000-C000-000000000046}"
Private Const IID_IOleObject            As String = "{00000112-0000-0000-C000-000000000046}"
Private Const IID_IOleInPlaceSite       As String = "{00000119-0000-0000-C000-000000000046}"
Private Const IID_IOleControlSite       As String = "{B196B289-BAB4-101A-B69C-00AA00341D07}"
Private ptrMe As Long

Private Const GMEM_FIXED As Long = &H0
Private Const asmPUSH_imm32 As Byte = &H68
Private Const asmRET_imm16 As Byte = &HC2
Private Const asmCALL_rel32 As Byte = &HE8

'*************************************************************************************************
' == Callback Subclassing by Paul Caton  =================================
' Local variables/constants: must declare these regardless if using subclassing, hooking, callbacks
Private z_scFunk            As Collection   'hWnd/thunk-address collection; initialized as needed
Private z_cbFunk            As Collection   'callback/thunk-address collection; initialized as needed
Private Const IDX_INDEX     As Long = 2     'index of the subclassed hWnd OR hook type
Private Const IDX_PREVPROC  As Long = 9     'Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11    'Thunk data index of the Before table for messages
Private Const IDX_ATABLE    As Long = 12    'Thunk data index of the After table for messages
Private Const IDX_CALLBACKORDINAL As Long = 36 ' Ubound(callback thunkdata)+1, index of the callback

' Declarations:
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Enum eThunkType
    SubclassThunk = 0
    CallbackThunk = 2
End Enum

'-Selfsub specific declarations----------------------------------------------------------------------------
Private Enum eMsgWhen                                                   'When to callback
  MSG_BEFORE = 1                                                        'Callback before the original WndProc
  MSG_AFTER = 2                                                         'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                            'Callback before and after the original WndProc
End Enum

' see ssc_Subclass for complete listing of indexes and what they relate to
Private Const IDX_PARM_USER As Long = 13    'Thunk data index of the User-defined callback parameter data index
Private Const IDX_UNICODE   As Long = 107   'Must be UBound(subclass thunkdata)+1; index for unicode support
Private Const MSG_ENTRIES   As Long = 32    'Number of msg table entries. Set to 1 if using ALL_MESSAGES for all subclassed windows
Private Const ALL_MESSAGES  As Long = -1    'All messages will callback

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'-------------------------------------------------------------------------------------------------

Private Function ssc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True, _
                    Optional ByRef bUnicode As Boolean = False, _
                    Optional ByVal bIsAPIwindow As Boolean = False) As Boolean 'Subclass the specified window handle

    '*************************************************************************************************
    '* lng_hWnd   - Handle of the window to subclass
    '* lParamUser - Optional, user-defined callback parameter
    '* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
    '* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety - Optional, enable/disable IDE safety measures. There is not reason to set this to False
    '* bUnicode - Optional, if True, Unicode API calls should be made to the window vs ANSI calls
    '*            Parameter is byRef and its return value should be checked to know if ANSI to be used or not
    '* bIsAPIwindow - Optional, if True DestroyWindow will be called if IDE ENDs
    '*****************************************************************************************
    '** Subclass.asm - subclassing thunk
    '**
    '** Paul_Caton@hotmail.com
    '** Copyright free, use and abuse as you see fit.
    '**
    '** v2.0 Re-write by LaVolpe, based mostly on Paul Caton's original thunks....... 20070720
    '** .... Reorganized & provided following additional logic
    '** ....... Unsubclassing only occurs after thunk is no longer recursed
    '** ....... Flag used to bypass callbacks until unsubclassing can occur
    '** ....... Timer used as delay mechanism to free thunk memory afer unsubclassing occurs
    '** .............. Prevents crash when one window subclassed multiple times
    '** .............. More END safe, even if END occurs within the subclass procedure
    '** ....... Added ability to destroy API windows when IDE terminates
    '** ....... Added auto-unsubclass when WM_NCDESTROY received
    '*****************************************************************************************
    ' Subclassing procedure must be declared identical to the one at the end of this class (Sample at Ordinal #1)

    Dim z_Sc(0 To IDX_UNICODE) As Long                 'Thunk machine-code initialised here
    
    Const SUB_NAME      As String = "ssc_Subclass"     'This routine's name
    Const CODE_LEN      As Long = 4 * IDX_UNICODE + 4  'Thunk length in bytes
    Const PAGE_RWX      As Long = &H40&                'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&              'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&              'Release allocated memory flag
    Const GWL_WNDPROC   As Long = -4                   'SetWindowsLong WndProc index
    Const WNDPROC_OFF   As Long = &H60                 'Thunk offset to the WndProc execution address
    Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1)) 'Bytes to allocate per thunk, data + code + msg tables
    
  ' This is the complete listing of thunk offset values and what they point/relate to.
  ' Those rem'd out are used elsewhere or are initialized in Declarations section
  
  'Const IDX_RECURSION  As Long = 0     'Thunk data index of callback recursion count
  'Const IDX_SHUTDOWN   As Long = 1     'Thunk data index of the termination flag
  'Const IDX_INDEX      As Long = 2     'Thunk data index of the subclassed hWnd
   Const IDX_EBMODE     As Long = 3     'Thunk data index of the EbMode function address
   Const IDX_CWP        As Long = 4     'Thunk data index of the CallWindowProc function address
   Const IDX_SWL        As Long = 5     'Thunk data index of the SetWindowsLong function address
   Const IDX_FREE       As Long = 6     'Thunk data index of the VirtualFree function address
   Const IDX_BADPTR     As Long = 7     'Thunk data index of the IsBadCodePtr function address
   Const IDX_OWNER      As Long = 8     'Thunk data index of the Owner object's vTable address
  'Const IDX_PREVPROC   As Long = 9     'Thunk data index of the original WndProc
   Const IDX_CALLBACK   As Long = 10    'Thunk data index of the callback method address
  'Const IDX_BTABLE     As Long = 11    'Thunk data index of the Before table
  'Const IDX_ATABLE     As Long = 12    'Thunk data index of the After table
  'Const IDX_PARM_USER  As Long = 13    'Thunk data index of the User-defined callback parameter data index
   Const IDX_DW         As Long = 14    'Thunk data index of the DestroyWinodw function address
   Const IDX_ST         As Long = 15    'Thunk data index of the SetTimer function address
   Const IDX_KT         As Long = 16    'Thunk data index of the KillTimer function address
   Const IDX_EBX_TMR    As Long = 20    'Thunk code patch index of the thunk data for the delay timer
   Const IDX_EBX        As Long = 26    'Thunk code patch index of the thunk data
  'Const IDX_UNICODE    As Long = xx    'Must be UBound(subclass thunkdata)+1; index for unicode support
    
    Dim z_ScMem       As Long           'Thunk base address
    Dim nAddr         As Long
    Dim nID           As Long
    Dim nMyID         As Long
    Dim bIDE          As Boolean

    If IsWindow(lng_hWnd) = 0 Then      'Ensure the window handle is valid
        Call zError(SUB_NAME, "Invalid window handle")
        Exit Function
    End If
    
    nMyID = GetCurrentProcessId                         'Get this process's ID
    GetWindowThreadProcessId lng_hWnd, nID              'Get the process ID associated with the window handle
    If nID <> nMyID Then                                'Ensure that the window handle doesn't belong to another process
        Call zError(SUB_NAME, "Window handle belongs to another process")
        Exit Function
    End If
    
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    
    nAddr = zAddressOf(oCallback, nOrdinal)             'Get the address of the specified ordinal method
    If nAddr = 0 Then                                   'Ensure that we've found the ordinal method
        Call zError(SUB_NAME, "Callback method not found")
        Exit Function
    End If
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    
    If z_ScMem <> 0 Then                                'Ensure the allocation succeeded
    
      If z_scFunk Is Nothing Then Set z_scFunk = New Collection 'If this is the first time through, do the one-time initialization
      On Error GoTo CatchDoubleSub                              'Catch double subclassing
      Call z_scFunk.Add(z_ScMem, "h" & lng_hWnd)                'Add the hWnd/thunk-address to the collection
      On Error GoTo 0
      
   'z_Sc (0) thru z_Sc(17) are used as storage for the thunks & IDX_ constants above relate to these thunk positions which are filled in below
    z_Sc(18) = &HD231C031: z_Sc(19) = &HBBE58960: z_Sc(21) = &H21E8F631: z_Sc(22) = &HE9000001: z_Sc(23) = &H12C&: z_Sc(24) = &HD231C031: z_Sc(25) = &HBBE58960: z_Sc(27) = &H3FFF631: z_Sc(28) = &H75047339: z_Sc(29) = &H2873FF23: z_Sc(30) = &H751C53FF: z_Sc(31) = &HC433913: z_Sc(32) = &H53FF2274: z_Sc(33) = &H13D0C: z_Sc(34) = &H18740000: z_Sc(35) = &H875C085: z_Sc(36) = &H820443C7: z_Sc(37) = &H90000000: z_Sc(38) = &H87E8&: z_Sc(39) = &H22E900: z_Sc(40) = &H90900000: z_Sc(41) = &H2C7B8B4A: z_Sc(42) = &HE81C7589: z_Sc(43) = &H90&: z_Sc(44) = &H75147539: z_Sc(45) = &H6AE80F: z_Sc(46) = &HD2310000: z_Sc(47) = &HE8307B8B: z_Sc(48) = &H7C&: z_Sc(49) = &H7D810BFF: z_Sc(50) = &H8228&: z_Sc(51) = &HC7097500: z_Sc(52) = &H80000443: z_Sc(53) = &H90900000: z_Sc(54) = &H44753339: z_Sc(55) = &H74047339: z_Sc(56) = &H2473FF3F: z_Sc(57) = &HFFFFFC68
    z_Sc(58) = &H2475FFFF: z_Sc(59) = &H811453FF: z_Sc(60) = &H82047B: z_Sc(61) = &HC750000: z_Sc(62) = &H74387339: z_Sc(63) = &H2475FF07: z_Sc(64) = &H903853FF: z_Sc(65) = &H81445B89: z_Sc(66) = &H484443: z_Sc(67) = &H73FF0000: z_Sc(68) = &H646844: z_Sc(69) = &H56560000: z_Sc(70) = &H893C53FF: z_Sc(71) = &H90904443: z_Sc(72) = &H10C261: z_Sc(73) = &H53E8&: z_Sc(74) = &H3075FF00: z_Sc(75) = &HFF2C75FF: z_Sc(76) = &H75FF2875: z_Sc(77) = &H2473FF24: z_Sc(78) = &H891053FF: z_Sc(79) = &H90C31C45: z_Sc(80) = &H34E30F8B: z_Sc(81) = &H1078C985: z_Sc(82) = &H4C781: z_Sc(83) = &H458B0000: z_Sc(84) = &H75AFF228: z_Sc(85) = &H90909023: z_Sc(86) = &H8D144D8D: z_Sc(87) = &H8D503443: z_Sc(88) = &H75FF1C45: z_Sc(89) = &H2C75FF30: z_Sc(90) = &HFF2875FF: z_Sc(91) = &H51502475: z_Sc(92) = &H2073FF52: z_Sc(93) = &H902853FF: z_Sc(94) = &H909090C3: z_Sc(95) = &H74447339: z_Sc(96) = &H4473FFF7
    z_Sc(97) = &H4053FF56: z_Sc(98) = &HC3447389: z_Sc(99) = &H89285D89: z_Sc(100) = &H45C72C75: z_Sc(101) = &H800030: z_Sc(102) = &H20458B00: z_Sc(103) = &H89145D89: z_Sc(104) = &H81612445: z_Sc(105) = &H4C4&: z_Sc(106) = &H1862FF00

    ' cache callback related pointers & offsets
      z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
      z_Sc(IDX_EBX_TMR) = z_ScMem                                             'Patch the thunk data address
      z_Sc(IDX_INDEX) = lng_hWnd                                              'Store the window handle in the thunk data
      z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
      z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
      z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
      z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
      z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
      
      ' validate unicode request & cache unicode usage
      If bUnicode Then bUnicode = (IsWindowUnicode(lng_hWnd) <> 0&)
      z_Sc(IDX_UNICODE) = bUnicode                                            'Store whether the window is using unicode calls or not
      
      ' get function pointers for the thunk
      If bIdeSafety = True Then                                               'If the user wants IDE protection
          Debug.Assert zInIDE(bIDE)
          If bIDE = True Then z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode", bUnicode) 'Store the EbMode function address in the thunk data
                                                        '^^ vb5 users, change vba6 to vba5
      End If
      If bIsAPIwindow Then                                                    'If user wants DestroyWindow sent should IDE end
          z_Sc(IDX_DW) = zFnAddr("user32", "DestroyWindow", bUnicode)
      End If
      z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree", bUnicode)           'Store the VirtualFree function address in the thunk data
      z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", bUnicode)        'Store the IsBadCodePtr function address in the thunk data
      z_Sc(IDX_ST) = zFnAddr("user32", "SetTimer", bUnicode)                  'Store the SetTimer function address in the thunk data
      z_Sc(IDX_KT) = zFnAddr("user32", "KillTimer", bUnicode)                 'Store the KillTimer function address in the thunk data
      
      If bUnicode Then
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcW", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongW", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongW(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      Else
          z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA", bUnicode)      'Store CallWindowProc function address in the thunk data
          z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA", bUnicode)       'Store the SetWindowLong function address in the thunk data
          RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                    'Copy the thunk code/data to the allocated memory
          z_Sc(IDX_PREVPROC) = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Set the new WndProc, return the address of the original WndProc
      End If
      If z_Sc(IDX_PREVPROC) = 0 Then                                          'Ensure the new WndProc was set correctly
          zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
          GoTo ReleaseMemory
      End If
      'Store the original WndProc address in the thunk data
      Call RtlMoveMemory(z_ScMem + IDX_PREVPROC * 4, VarPtr(z_Sc(IDX_PREVPROC)), 4&)
      ssc_Subclass = True                                                     'Indicate success
    Else
        Call zError(SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError)
    End If
 Exit Function                                                                'Exit ssc_Subclass
    
CatchDoubleSub:
 Call zError(SUB_NAME, "Window handle is already subclassed")
      
ReleaseMemory:
      Call VirtualFree(z_ScMem, 0, MEM_RELEASE)                               'ssc_Subclass has failed after memory allocation, so release the memory
End Function

'Terminate all subclassing
Private Sub ssc_Terminate()
    ' can be made public, can be removed & zTerminateThunks can be called instead
    Call zTerminateThunks(SubclassThunk)
End Sub

'UnSubclass the specified window handle
Private Sub ssc_UnSubclass(ByVal lng_hWnd As Long)
    ' can be made public, can be removed & zUnthunk can be called instead
    Call zUnThunk(lng_hWnd, SubclassThunk)
End Sub

'Add the message value to the window handle's specified callback table
Private Sub ssc_AddMsg(ByVal lng_hWnd As Long, ByVal When As eMsgWhen, ParamArray Messages() As Variant)
    Dim z_ScMem       As Long                                   'Thunk base address
    
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)           'Ensure that the thunk hasn't already released its memory
    If z_ScMem Then
      Dim M As Long
      For M = LBound(Messages) To UBound(Messages)
        Select Case VarType(Messages(M))                        ' ensure no strings, arrays, doubles, objects, etc are passed
        Case vbByte, vbInteger, vbLong
            If When And MSG_BEFORE Then                         'If the message is to be added to the before original WndProc table...
              If zAddMsg(Messages(M), IDX_BTABLE, z_ScMem) = False Then 'Add the message to the before table
                When = (When And Not MSG_BEFORE)
              End If
            End If
            If When And MSG_AFTER Then                          'If message is to be added to the after original WndProc table...
              If zAddMsg(Messages(M), IDX_ATABLE, z_ScMem) = False Then 'Add the message to the after table
                When = (When And Not MSG_AFTER)
              End If
            End If
        End Select
      Next
    End If
End Sub

'Call the original WndProc
Private Function ssc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ' can be made public, can be removed if you will not use this in your window procedure
    Dim z_ScMem       As Long                           'Thunk base address
    z_ScMem = zMap_VFunction(lng_hWnd, SubclassThunk)
    If z_ScMem Then                                     'Ensure that the thunk hasn't already released its memory
        If zData(IDX_UNICODE, z_ScMem) Then
            ssc_CallOrigWndProc = CallWindowProcW(zData(IDX_PREVPROC, z_ScMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        Else
            ssc_CallOrigWndProc = CallWindowProcA(zData(IDX_PREVPROC, z_ScMem), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
        End If
    End If
End Function
'Add the message to the specified table of the window handle
Private Function zAddMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long) As Boolean
      Dim nCount As Long                            'Table entry count
      Dim nBase  As Long
      Dim i      As Long                            'Loop index
    
      zAddMsg = True
      nBase = zData(nTable, z_ScMem)                'Map zData() to the specified table
      
      If uMsg = ALL_MESSAGES Then                   'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                       'Set the table entry count to ALL_MESSAGES
      Else
        
        nCount = zData(0, nBase)                    'Get the current table entry count
        For i = 1 To nCount                         'Loop through the table entries
          If zData(i, nBase) = 0 Then               'If the element is free...
            zData(i, nBase) = uMsg                  'Use this element
            GoTo Bail                               'Bail
          ElseIf zData(i, nBase) = uMsg Then        'If the message is already in the table...
            GoTo Bail                               'Bail
          End If
        Next i                                      'Next message table entry
    
        nCount = i                                  'On drop through: i = nCount + 1, the new table entry count
        If nCount > MSG_ENTRIES Then                'Check for message table overflow
          Call zError("zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values")
          zAddMsg = False
          GoTo Bail
        End If
        
        zData(nCount, nBase) = uMsg                                            'Store the message in the appended table entry
      End If
    
      zData(0, nBase) = nCount                                                 'Store the new table entry count
Bail:
End Function

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long, ByVal z_ScMem As Long)
      Dim nCount As Long                                                        'Table entry count
      Dim nBase  As Long
      Dim i      As Long                                                        'Loop index
    
      nBase = zData(nTable, z_ScMem)                                            'Map zData() to the specified table
    
      If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
        zData(0, nBase) = 0                                                     'Zero the table entry count
      Else
        nCount = zData(0, nBase)                                                'Get the table entry count
        
        For i = 1 To nCount                                                     'Loop through the table entries
          If zData(i, nBase) = uMsg Then                                        'If the message is found...
            zData(i, nBase) = 0                                                 'Null the msg value -- also frees the element for re-use
            GoTo Bail                                                           'Bail
          End If
        Next i                                                                  'Next message table entry
       ' zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
      End If
Bail:
End Sub

'-The following routines are exclusively for the scb_SetCallbackAddr routines----------------------------
Public Function scb_SetCallbackAddr(ByVal nParamCount As Long, _
                     Optional ByVal nOrdinal As Long = 1, _
                     Optional ByVal oCallback As Object = Nothing, _
                     Optional ByVal bIdeSafety As Boolean = True, _
                     Optional ByVal bIsTimerCallback As Boolean) As Long   'Return the address of the specified callback thunk
    '*************************************************************************************************
    '* nParamCount  - The number of parameters that will callback
    '* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
    '* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
    '* bIdeSafety   - Optional, set to false to disable IDE protection.
    '* bIsTimerCallback - optional, set to true for extra protection when used as a SetTimer callback
    '       If True, timer will be destroyed when IDE/app terminates. See scb_ReleaseCallback.
    '*************************************************************************************************
    ' Callback procedure must return a Long even if, per MSDN, the callback procedure is a Sub vs Function
    ' The number of parameters and their types are dependent on the individual callback procedures
    
    Const MEM_LEN     As Long = IDX_CALLBACKORDINAL * 4 + 4     'Memory bytes required for the callback thunk
    Const PAGE_RWX    As Long = &H40&                           'Allocate executable memory
    Const MEM_COMMIT  As Long = &H1000&                         'Commit allocated memory
    Const SUB_NAME      As String = "scb_SetCallbackAddr"       'This routine's name
    Const INDX_OWNER    As Long = 0                             'Thunk data index of the Owner object's vTable address
    Const INDX_CALLBACK As Long = 1                             'Thunk data index of the EbMode function address
    Const INDX_EBMODE   As Long = 2                             'Thunk data index of the IsBadCodePtr function address
    Const INDX_BADPTR   As Long = 3                             'Thunk data index of the IsBadCodePtr function address
    Const INDX_KT       As Long = 4                             'Thunk data index of the KillTimer function address
    Const INDX_EBX      As Long = 6                             'Thunk code patch index of the thunk data
    Const INDX_PARAMS   As Long = 18                            'Thunk code patch index of the number of parameters expected in callback
    Const INDX_PARAMLEN As Long = 24                            'Thunk code patch index of the bytes to be released after callback
    Const PROC_OFF      As Long = &H14                          'Thunk offset to the callback execution address

    Dim z_ScMem       As Long                                   'Thunk base address
    Dim z_Cb()    As Long                                       'Callback thunk array
    Dim nValue    As Long
    Dim nCallback As Long
    Dim bIDE      As Boolean
      
    If oCallback Is Nothing Then Set oCallback = Me     'If the user hasn't specified the callback owner
    If z_cbFunk Is Nothing Then
        Set z_cbFunk = New Collection           'If this is the first time through, do the one-time initialization
    Else
        On Error Resume Next                    'Catch already initialized?
        z_ScMem = z_cbFunk.Item("h" & ObjPtr(oCallback) & "." & nOrdinal) 'Test it
        If Err = 0 Then
            scb_SetCallbackAddr = z_ScMem + PROC_OFF  'we had this one, just reference it
            Exit Function
        End If
        On Error GoTo 0
    End If
    
    If nParamCount < 0 Then                     ' validate parameters
        Call zError(SUB_NAME, "Invalid Parameter count")
        Exit Function
    End If
    
    nCallback = zAddressOf(oCallback, nOrdinal)         'Get the callback address of the specified ordinal
    If nCallback = 0 Then
        Call zError(SUB_NAME, "Callback address not found.")
        Exit Function
    End If
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
        
    If z_ScMem = 0& Then
        Call zError(SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError)  ' oops
        Exit Function
    End If
    Call z_cbFunk.Add(z_ScMem, "h" & ObjPtr(oCallback) & "." & nOrdinal) 'Add the callback/thunk-address to the collection
        
    ReDim z_Cb(0 To IDX_CALLBACKORDINAL) As Long          'Allocate for the machine-code array
    
    ' Create machine-code array
    z_Cb(5) = &HBB60E089: z_Cb(7) = &H73FFC589: z_Cb(8) = &HC53FF04: z_Cb(9) = &H59E80A74: z_Cb(10) = &HE9000000
    z_Cb(11) = &H30&: z_Cb(12) = &H87B81: z_Cb(13) = &H75000000: z_Cb(14) = &H9090902B: z_Cb(15) = &H42DE889: z_Cb(16) = &H50000000: z_Cb(17) = &HB9909090: z_Cb(19) = &H90900AE3
    z_Cb(20) = &H8D74FF: z_Cb(21) = &H9090FAE2: z_Cb(22) = &H53FF33FF: z_Cb(23) = &H90909004: z_Cb(24) = &H2BADC261: z_Cb(25) = &H3D0853FF: z_Cb(26) = &H1&: z_Cb(27) = &H23DCE74: z_Cb(28) = &H74000000: z_Cb(29) = &HAE807
    z_Cb(30) = &H90900000: z_Cb(31) = &H4589C031: z_Cb(32) = &H90DDEBFC: z_Cb(33) = &HFF0C75FF: z_Cb(34) = &H53FF0475: z_Cb(35) = &HC310&

    z_Cb(INDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr", False)
    z_Cb(INDX_OWNER) = ObjPtr(oCallback)                    'Set the Owner
    z_Cb(INDX_CALLBACK) = nCallback                         'Set the callback address
    z_Cb(IDX_CALLBACKORDINAL) = nOrdinal                    'Cache ordinal used for zTerminateThunks
      
    If bIdeSafety = True Then                               'If the user wants IDE protection
        Debug.Assert zInIDE(bIDE)
        If bIDE = True Then z_Cb(INDX_EBMODE) = zFnAddr("vba6", "EbMode", False) 'Store the EbMode function address in the thunk data
    End If
    If bIsTimerCallback Then
        z_Cb(INDX_KT) = zFnAddr("user32", "KillTimer", False)
    End If
        
    z_Cb(INDX_PARAMS) = nParamCount                         'Set the parameter count
    Call RtlMoveMemory(VarPtr(z_Cb(INDX_PARAMLEN)) + 2, VarPtr(nParamCount * 4), 2&)

    z_Cb(INDX_EBX) = z_ScMem                                'Set the data address relative to virtual memory pointer

    Call RtlMoveMemory(z_ScMem, VarPtr(z_Cb(INDX_OWNER)), MEM_LEN) 'Copy thunk code to executable memory
    scb_SetCallbackAddr = z_ScMem + PROC_OFF                       'Thunk code start address
End Function

Public Sub scb_ReleaseCallback(ByVal nOrdinal As Long, Optional ByVal oCallback As Object)
    ' can be made public, can be removed & zUnThunk can be called instead
    ' NEVER call this from within the callback routine itself
    
    ' oCallBack is the object containing nOrdinal to be released
    ' if oCallback was already closed (say it was a class or form), then you won't be
    '   able to release it here, but it will be released when zTerminateThunks is
    '   eventually called
    
    ' Special Warning. If the callback thunk is used for a recurring callback (i.e., Timer),
    ' then ensure you terminate what is using the callback before releasing the thunk,
    ' otherwise you are subject to a crash when that item tries to callback to zeroed memory
    Call zUnThunk(nOrdinal, CallbackThunk, oCallback)
End Sub

Public Sub scb_TerminateCallbacks()
    ' can be made public, can be removed & zTerminateThunks can be called instead
    Call zTerminateThunks(CallbackThunk)
End Sub

'========================================================================
' COMMON USE ROUTINES
'-The following routines are used for each of the three types of thunks
'========================================================================

'-The following routines are used for each of the three types of thunks ----------------------------

'Maps zData() to the memory address for the specified thunk type
Private Function zMap_VFunction(vFuncTarget As Long, _
                                vType As eThunkType, _
                                Optional oCallback As Object, _
                                Optional bIgnoreErrors As Boolean) As Long
    
    Dim thunkCol As Collection
    Dim colID As String
    Dim z_ScMem       As Long         'Thunk base address
    
    If vType = CallbackThunk Then
        Set thunkCol = z_cbFunk
        If oCallback Is Nothing Then Set oCallback = Me
        colID = "h" & ObjPtr(oCallback) & "." & vFuncTarget
    ElseIf vType = SubclassThunk Then
        Set thunkCol = z_scFunk
        colID = "h" & vFuncTarget
    Else
        Call zError("zMap_Vfunction", "Invalid thunk type passed")
        Exit Function
    End If
    
    If thunkCol Is Nothing Then
        Call zError("zMap_VFunction", "Thunk hasn't been initialized")
    Else
        If thunkCol.Count Then
            On Error GoTo Catch
            z_ScMem = thunkCol(colID)               'Get the thunk address
            If IsBadCodePtr(z_ScMem) Then z_ScMem = 0&
            zMap_VFunction = z_ScMem
        End If
    End If
    Exit Function                                   'Exit returning the thunk address
Catch:
    ' error ignored when zUnThunk is called, error handled there
    If Not bIgnoreErrors Then Call zError("zMap_VFunction", "Thunk type for " & vType & " does not exist")
End Function

' sets/retrieves data at the specified offset for the specified memory address
Private Property Get zData(ByVal nIndex As Long, ByVal z_ScMem As Long) As Long
  Call RtlMoveMemory(VarPtr(zData), z_ScMem + (nIndex * 4), 4)
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal z_ScMem As Long, ByVal nValue As Long)
  Call RtlMoveMemory(z_ScMem + (nIndex * 4), VarPtr(nValue), 4)
End Property

'Error handler
Private Sub zError(ByRef sRoutine As String, ByVal sMsg As String)
  ' Note. These two lines can be rem'd out if you so desire. But don't remove the routine
  ' App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  Call MsgBox(sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine)
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String, ByVal asUnicode As Boolean) As Long
  If asUnicode Then
    zFnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)         'Get the specified procedure address
  Else
    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                 'Get the specified procedure address
  End If
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
  ' ^^ FYI VB5 users. Search for zFnAddr("vba6", "EbMode") and replace with zFnAddr("vba5", "EbMode")
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
    ' Note: used both in subclassing and hooking routines
  Dim bSub  As Byte                                     'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                     'Address of the vTable
  Dim i     As Long                                     'Loop index
  Dim J     As Long                                     'Loop limit
  
  Call RtlMoveMemory(VarPtr(nAddr), ObjPtr(oCallback), 4) 'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then             'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then          'Probe for a Form method
      If Not zProbe(nAddr + &H710, i, bSub) Then        'Probe for a PropertyPage method
        If Not zProbe(nAddr + &H7A4, i, bSub) Then      'Probe for a UserControl method
            Exit Function                               'Bail...
        End If
      End If
    End If
  End If
  
  i = i + 4                                             'Bump to the next entry
  J = i + 1024                                          'Set a reasonable limit, scan 256 vTable entries
  Do While i < J
    Call RtlMoveMemory(VarPtr(nAddr), i, 4)             'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                     'Is the entry an invalid code address?
      Call RtlMoveMemory(VarPtr(zAddressOf), i - (nOrdinal * 4), 4) 'Return the specified vTable entry address
      Exit Do                                                       'Bad method signature, quit loop
    End If

    Call RtlMoveMemory(VarPtr(bVal), nAddr, 1)                      'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                            'If the byte doesn't match the expected value...
      Call RtlMoveMemory(VarPtr(zAddressOf), i - (nOrdinal * 4), 4) 'Return the specified vTable entry address
      Exit Do                                                       'Bad method signature, quit loop
    End If
    
    i = i + 4                                                       'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                    'Start address
  nLimit = nAddr + 32                                               'Probe eight entries
  Do While nAddr < nLimit                                           'While we've not reached our probe depth
    Call RtlMoveMemory(VarPtr(nEntry), nAddr, 4)                    'Get the vTable entry
    
    If nEntry <> 0 Then                                             'If not an implemented interface
      Call RtlMoveMemory(VarPtr(bVal), nEntry, 1)                   'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                            'Check for a native or pcode method signature
        nMethod = nAddr                                             'Store the vTable entry
        bSub = bVal                                                 'Store the found method signature
        zProbe = True                                               'Indicate success
        Exit Do                                                     'Return
      End If
    End If
    nAddr = nAddr + 4                                               'Next vTable entry
  Loop
End Function

Private Function zInIDE(ByRef bIDE As Boolean) As Boolean
    ' only called in IDE, never called when compiled
    bIDE = True
    zInIDE = bIDE
End Function

Private Sub zUnThunk(ByVal thunkID As Long, ByVal vType As eThunkType, Optional ByVal oCallback As Object)
    ' thunkID, depends on vType:
    '   - Subclassing:  the hWnd of the window subclassed
    '   - Callbacks:    the ordinal of the callback
    '       ensure KillTimer is already called, if any callback used for SetTimer
    ' oCallback only used when vType is CallbackThunk
    Const IDX_SHUTDOWN  As Long = 1
    Const MEM_RELEASE As Long = &H8000&             'Release allocated memory flag
    
    Dim z_ScMem       As Long                       'Thunk base address
    
    z_ScMem = zMap_VFunction(thunkID, vType, oCallback, True)
    Select Case vType
    Case SubclassThunk
        If z_ScMem Then                         'Ensure that the thunk hasn't already released its memory
            zData(IDX_SHUTDOWN, z_ScMem) = 1                  'Set the shutdown indicator
            Call zDelMsg(ALL_MESSAGES, IDX_BTABLE, z_ScMem)   'Delete all before messages
            Call zDelMsg(ALL_MESSAGES, IDX_ATABLE, z_ScMem)   'Delete all after messages
        End If
        Call z_scFunk.Remove("h" & thunkID)                   'Remove the specified thunk from the collection
    Case CallbackThunk
        If z_ScMem Then                         'Ensure that the thunk hasn't already released its memory
            Call VirtualFree(z_ScMem, 0, MEM_RELEASE)   'Release allocated memory
        End If
        Call z_cbFunk.Remove("h" & ObjPtr(oCallback) & "." & thunkID) 'Remove the specified thunk from the collection
    End Select
End Sub

Private Sub zTerminateThunks(ByVal vType As eThunkType)
    ' Terminates all thunks of a specific type
    ' Any subclassing, recurring callbacks should have already been canceled
    Dim i As Long
    Dim oCallback As Object
    Dim thunkCol As Collection
    Dim z_ScMem       As Long                           'Thunk base address
    Const INDX_OWNER As Long = 0
    
    Select Case vType
    Case SubclassThunk
        Set thunkCol = z_scFunk
    Case CallbackThunk
        Set thunkCol = z_cbFunk
    Case Else
        Exit Sub
    End Select
    
    If Not (thunkCol Is Nothing) Then                 'Ensure that hooking has been started
      With thunkCol
        For i = .Count To 1 Step -1                   'Loop through the collection of hook types in reverse order
          z_ScMem = .Item(i)                          'Get the thunk address
          If IsBadCodePtr(z_ScMem) = 0 Then           'Ensure that the thunk hasn't already released its memory
            Select Case vType
                Case SubclassThunk
                    zUnThunk zData(IDX_INDEX, z_ScMem), SubclassThunk    'Unsubclass
                Case CallbackThunk
                    ' zUnThunk expects object not pointer, convert pointer to object
                    Call RtlMoveMemory(VarPtr(oCallback), VarPtr(zData(INDX_OWNER, z_ScMem)), 4&)
                    Call zUnThunk(zData(IDX_CALLBACKORDINAL, z_ScMem), CallbackThunk, oCallback) ' release callback
                    ' remove the object pointer reference
                    Call RtlMoveMemory(VarPtr(oCallback), VarPtr(INDX_OWNER), 4&)
            End Select
          End If
        Next i                                        'Next member of the collection
      End With
      Set thunkCol = Nothing                         'Destroy the hook/thunk-address collection
    End If
End Sub
'=========================================================================

' === UserControl Events =================================================
Private Sub UserControl_Click()
     If IsInside Then RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    If IsInside Then RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If IsInside Then RaiseEvent MouseDown(Button, Shift, X, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If IsInside Then RaiseEvent MouseMove(Button, Shift, X, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If IsInside Then RaiseEvent MouseUp(Button, Shift, X, y)
End Sub

Private Sub UserControl_GotFocus()
    Call SetFocus(hTabs)
End Sub

Private Sub UserControl_Initialize()
    hMod = LoadLibraryA("shell32.dlL")
    Call InitCommonControls
    Set m_oFont = New StdFont
End Sub

Private Sub UserControl_InitProperties()
    Call CrearTabStrip
    Set Font = Ambient.Font
    Call AddTab(0, "Tab 0")
    Call AddTab(1, "Tab 1")
    Call AddTab(2, "Tab 2")
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If hTabs = 0 Then Call CrearTabStrip
    With PropBag
        Set Font = .ReadProperty("Font", Ambient.Font)
    End With
    Call pvReadControls(PropBag)
    Call pvUpdateTabView
    If Ambient.UserMode Then
        If ssc_Subclass(hWnd) Then
            Call ssc_AddMsg(hWnd, MSG_BEFORE, WM_NOTIFY, WM_MOUSEACTIVATE)
        End If
        If ssc_Subclass(hTabs) Then
            Call ssc_AddMsg(hTabs, MSG_BEFORE, WM_SETFOCUS, WM_DESTROY, WM_KEYDOWN, WM_CHAR, WM_KEYUP, WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN, WM_MOUSEMOVE, WM_MOUSELEAVE, WM_LBUTTONUP, WM_RBUTTONUP, WM_MBUTTONUP, TCM_DELETEALLITEMS)
        End If
        Call pvInitIPAO
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim s As String, i As Long
    For i = 0 To Count - 1
        If i > 0 Then s = s & m_Splitter
        s = s & ItemText(i)
    Next
    With PropBag
        Call .WriteProperty("Font", Font, Ambient.Font)
        Call .WriteProperty("Tabs", s, vbNullString)
    End With
    Call pvWriteControls(PropBag)
End Sub

Private Sub UserControl_Resize()
    If hTabs Then Call SetWindowPos(hTabs, 0, 0, 0, ScaleWidth, ScaleHeight, SWP_NOZORDER Or SWP_NOOWNERZORDER)
End Sub

Private Sub UserControl_Show()
    Call UserControl_Resize
    Call pvUpdateTabView
End Sub

Private Sub UserControl_Terminate()
    Call zTerminateThunks(CallbackThunk)
    Call zTerminateThunks(SubclassThunk)
    Call DestroyWindow(hTabs)
    Call FreeLibrary(hMod)
    Set m_oFont = Nothing
    Erase m_controlsTab
End Sub

' === Public Function ====================================================
Public Function AddTab(ByVal index As Long, ByVal Caption As String, Optional ByVal ItemData As Long, Optional ByVal ImageIndex As Long = -1) As Boolean
    If hTabs = 0 Then Exit Function
    Dim sTabSrip As TCITEM
    With sTabSrip
        .mask = TCIF_TEXT Or TCIF_IMAGE Or TCIF_PARAM
        .iImage = ImageIndex
        .lParam = ItemData
        .pszText = StrPtr(Caption)
    End With
    
    AddTab = SendMessageW(hTabs, TCM_INSERTITEMW, index, sTabSrip)
    Call PropertyChanged("Tabs")
End Function
Public Function Clear() As Boolean
    If hTabs Then Clear = SendMessageW(hTabs, TCM_DELETEALLITEMS, 0, ByVal 0)
    Call PropertyChanged("Tabs")
End Function
Public Function HitTest(ByVal X As Single, ByVal y As Single) As Long
    If hTabs Then
        Dim ht As TCHITTESTINFO
        ht.pt.X = ScaleX(X, ScaleMode, vbPixels)
        ht.pt.y = ScaleY(y, ScaleMode, vbPixels)
        ht.flags = TCHT_ONITEM
        HitTest = SendMessageW(hTabs, TCM_HITTEST, 0, ht)
    End If
End Function
Public Function RemoveTab(ByVal index As Long) As Boolean
    If hTabs Then RemoveTab = SendMessageW(hTabs, TCM_DELETEITEM, index, ByVal 0)
    Call PropertyChanged("Tabs")
End Function
Public Function SetImageList(ByVal hImageList As Long) As Boolean
    If hTabs Then SetImageList = SendMessageW(hTabs, TCM_SETIMAGELIST, 0, ByVal hImageList)
End Function
Public Function SetMinTabWidth(ByVal newMinWidth As Long) As Boolean
    If hTabs Then SetMinTabWidth = SendMessageW(hTabs, TCM_SETMINTABWIDTH, 0, ByVal newMinWidth)
End Function

Public Function ClientTop() As Long
    If hTabs = 0 Then Exit Function
    Dim r As RECT
    Call SendMessageW(hTabs, TCM_ADJUSTRECT, False, r)
    ClientTop = r.Top
End Function
Public Function ClientLeft() As Long
    If hTabs = 0 Then Exit Function
    Dim r As RECT
    Call SendMessageW(hTabs, TCM_ADJUSTRECT, False, r)
    ClientLeft = r.Left
End Function

Public Function ClientWidth() As Long
    If hTabs = 0 Then Exit Function
    Dim r As RECT
    Call SendMessageW(hTabs, TCM_ADJUSTRECT, False, r)
    ClientWidth = ScaleWidth + r.Right
End Function

Public Function ClientHeight() As Long
    If hTabs = 0 Then Exit Function
    Dim r As RECT
    Call SendMessageW(hTabs, TCM_ADJUSTRECT, False, r)
    ClientHeight = ScaleHeight + r.Bottom
End Function

' === Private Functions ==================================================
Private Function IsInside() As Boolean
    If hTabs = 0 Then Exit Function
    Dim r As RECT, p As POINT
    Call SendMessageW(hTabs, TCM_ADJUSTRECT, False, r)
    r.Right = ScaleWidth + r.Right
    r.Bottom = ScaleHeight + r.Bottom
    Call pvUCCoordPixel(p.X, p.y)
    IsInside = p.X >= r.Left And p.X <= r.Right And p.y >= r.Top And p.y <= r.Bottom
End Function
Private Sub CrearTabStrip()
    hTabs = CreateWindowExW(WS_EX_TABS, StrPtr(WC_TABCONTROL), vbNullString, WS_TABS, 0, 0, (UserControl.ScaleWidth), (UserControl.ScaleHeight), UserControl.hWnd, 0, App.hInstance, ByVal 0&)
End Sub

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    Set Font = m_oFont
End Sub

Private Function pvDestroyFont() As Boolean
    If (m_hFont) Then
        If (DeleteObject(m_hFont)) Then
            pvDestroyFont = True
            m_hFont = 0
        End If
    End If
End Function

Private Function pvShiftState() As Integer
  Dim lS As Integer
    If (GetAsyncKeyState(vbKeyShift) < 0) Then lS = lS Or vbShiftMask
    If (GetAsyncKeyState(vbKeyMenu) < 0) Then lS = lS Or vbAltMask
    If (GetAsyncKeyState(vbKeyControl) < 0) Then lS = lS Or vbCtrlMask
    pvShiftState = lS
End Function

Private Function pvButton(ByVal uMsg As Long) As Integer
    Select Case uMsg
        Case WM_LBUTTONDOWN, WM_LBUTTONUP
            pvButton = vbLeftButton
        Case WM_RBUTTONDOWN, WM_RBUTTONUP
            pvButton = vbRightButton
        Case WM_MBUTTONDOWN, WM_MBUTTONUP
            pvButton = vbMiddleButton
        Case WM_MOUSEMOVE
            Select Case True
                Case GetAsyncKeyState(vbKeyLButton) < 0
                    pvButton = vbLeftButton
                Case GetAsyncKeyState(vbKeyRButton) < 0
                    pvButton = vbRightButton
                Case GetAsyncKeyState(vbKeyMButton) < 0
                    pvButton = vbMiddleButton
            End Select
    End Select
End Function

Private Sub pvUCCoordPixel(X As Long, y As Long)
    Dim uPt As POINT
    Call GetCursorPos(uPt)
    Call ScreenToClient(hTabs, uPt)
    X = uPt.X
    y = uPt.y
End Sub

Private Sub pvUCCoordScale(X As Single, y As Single)
    Dim uPt As POINT
    Call GetCursorPos(uPt)
    Call ScreenToClient(hTabs, uPt)
    X = ScaleX(uPt.X, vbPixels, UserControl.ScaleMode)
    y = ScaleY(uPt.y, vbPixels, UserControl.ScaleMode)
End Sub

Private Sub pvTrackMouseLeave(ByVal lng_hWnd As Long)
    'Track the mouse leaving the indicated window
  Dim uTME As TRACKMOUSEEVENT_STRUCT
    With uTME
        .cbSize = Len(uTME)
        .dwFlags = TME_LEAVE
        .hwndTrack = lng_hWnd
    End With
    Call TrackMouseEvent(uTME)
End Sub

'==== Saving control related =======================
Private Function pvGetControlName(oCtl As Control) As String
    pvGetControlName = oCtl.Name
    If oCtl.index <> -1 Then pvGetControlName = pvGetControlName & "(" & oCtl.index
End Function

Private Function pvIsControl(oCtl As Control, sName As String) As Boolean
    If oCtl Is Nothing Then Exit Function
    'If Not (oCtl) Then Exit Function
    'If oCtl Then
        Dim i As Integer
        i = InStr(sName, "(")
        If i = 0 Then
            pvIsControl = oCtl.Name = sName
        Else
            pvIsControl = oCtl.Name = Mid$(sName, 1, i - 1) And oCtl.index = Mid$(sName, i + 1)
        End If
    'End If
End Function

Private Sub pvReadControls(PropBag As PropertyBag)
    Dim s As String, i As Long, tabs() As String, J As Long, k As Long, z As Long
    With PropBag
        s = .ReadProperty("Tabs", vbNullString)
        If s <> vbNullString Then
            Call Clear
            tabs = Split(s, m_Splitter)
            For i = LBound(tabs) To UBound(tabs)
                Call AddTab(i, tabs(i))
            Next
        End If
        If .ReadProperty("bSaved", False) Then
            z = .ReadProperty("ControlCount", 0)
            ReDim m_controlsTab(z)
            For J = 0 To z
                'tabs = Split(.ReadProperty("ControlsTab" & J), m_Splitter)
                'k = UBound(tabs)
                m_controlsTab(J).Name = .ReadProperty("CN" & J)
                m_controlsTab(J).TabIndex = .ReadProperty("CT" & J)
                'If k > 0 Then
'                    ReDim m_controlsTab(UBound(m_controlsTab) + k)
'                    For i = 0 To k
'                        m_controlsTab(i).TabIndex = j
'                        m_controlsTab(i).Name = tabs(i)
'                    Next
                'End If
            Next
        End If
    End With
End Sub

Private Sub pvWriteControls(PropBag As PropertyBag)
    Dim i As Long, J As Long, z As Long, y As Long
    With PropBag
    If ContainedControls.Count > 0 Then
        Call pvSaveControlState
        If (Not Not m_controlsTab) Then
            z = UBound(m_controlsTab)
            Call .WriteProperty("ControlCount", z)
            y = Count
            'For J = 0 To y
                's = vbNullString
                For i = 0 To z
                    'If LenB(m_controlsTab(i).Name) Then
                    '    If m_controlsTab(i).TabIndex = J Then
                            'If i > 0 Then s = s & m_Splitter
                            's = s & m_controlsTab(i).Name
                            'Debug.Print "Saved " & m_controlsTab(i).Name & " in Tab " & j
                            Call .WriteProperty("CN" & i, m_controlsTab(i).Name)
                            Call .WriteProperty("CT" & i, m_controlsTab(i).TabIndex)
                    '    End If
                    'End If
                Next
                'Call .WriteProperty("ControlsTab" & J, s)
            'Next
            Call .WriteProperty("bSaved", True)
        End If
    End If
    End With
End Sub

Private Sub pvSaveControlState()
    If Ambient.UserMode Then Exit Sub
    If UserControl.ContainedControls.Count > 0 Then
        Dim i As Long, lEnd As Long, oCtl As Control, J As Integer, k As Integer, z As Integer
        lEnd = ContainedControls.Count - 1
        If (Not Not m_controlsTab) = False Then
            ' First time using the Tab control
            ' Save all data :E
            ReDim Preserve m_controlsTab(lEnd)
            For Each oCtl In UserControl.ContainedControls
                m_controlsTab(i).Name = pvGetControlName(oCtl)
                m_controlsTab(i).TabIndex = SelectedItem
                i = i + 1
            Next
            Set oCtl = Nothing
        Else
            ' There are more controls Than Before?
            z = UBound(m_controlsTab)
            If lEnd > z Then ReDim Preserve m_controlsTab(lEnd)
            
            ' Delete old indexed controls form array
            For i = 0 To z
                For J = 0 To lEnd
                    If pvIsControl(ContainedControls(J), m_controlsTab(i).Name) Then Exit For
                Next
                If J > lEnd Then m_controlsTab(i).Name = vbNullString
            Next
            i = 0
            For Each oCtl In UserControl.ContainedControls
                For J = 0 To z
                    If pvIsControl(oCtl, m_controlsTab(J).Name) Then
                        ' The control existed
                        If IsWindowVisible(oCtl.hWnd) Then
                            m_controlsTab(J).TabIndex = SelectedItem
                        End If
                        Exit For
                    End If
                Next
                If J > z Then
                    For J = 0 To z
                        If LenB(m_controlsTab(J).Name) = 0 Then
                            m_controlsTab(J).Name = pvGetControlName(oCtl)
                            m_controlsTab(J).TabIndex = SelectedItem
                            Exit For
                        End If
                    Next
                End If
                i = i + 1
            Next
            Set oCtl = Nothing
        End If
'        z = UBound(m_controlsTab)
'        For i = 0 To z
'            Debug.Print m_controlsTab(i).Name & " in Tab " & m_controlsTab(i).TabIndex
'        Next
    End If
End Sub

Private Sub pvUpdateTabView()
    If (Not Not m_controlsTab) = False Then Exit Sub
    If UserControl.ContainedControls.Count > 0 Then
        Dim i As Long, oCtl As Control, k As Long, X As Long, z As Long
        k = UBound(m_controlsTab)
        z = ContainedControls.Count - 1
        For Each oCtl In ContainedControls
        'For X = 0 To z
            For i = 0 To k
                'oCtl = ContainedControls.Item(X)
                If pvIsControl(oCtl, m_controlsTab(i).Name) Then
                    If m_controlsTab(i).TabIndex = SelectedItem Then
                        Call ShowWindow(oCtl.hWnd, 1)
                    Else
                        Call ShowWindow(oCtl.hWnd, 0)
                    End If
                    Exit For
                End If
            Next
        Next
    End If
End Sub


' === Properties =========================================================
Public Property Get Align() As AlignConstants
    Align = vbAlignTop
    If hTabs Then
        Dim style As Long
        If style And (TCS_VERTICAL Or TCS_RIGHT) Then
            Align = vbAlignRight
        ElseIf style And TCS_VERTICAL Then
            Align = vbAlignLeft
        ElseIf style And TCS_BOTTOM Then
            Align = vbAlignBottom
        End If
    End If
End Property
Public Property Let Align(ByVal value As AlignConstants)
    If hTabs Then
        Dim style As Long
        style = GetWindowLongW(hTabs, GWL_STYLE) And Not (TCS_BOTTOM Or TCS_VERTICAL Or TCS_RIGHT)
        Select Case value
            Case AlignConstants.vbAlignBottom
                style = style Or TCS_BOTTOM
            Case AlignConstants.vbAlignLeft
                style = style Or TCS_VERTICAL
            Case AlignConstants.vbAlignRight
                style = style Or TCS_VERTICAL Or TCS_RIGHT
        End Select
        Call SetWindowLongW(hTabs, GWL_STYLE, style)
        Call PropertyChanged("Align")
    End If
End Property

Public Property Get Count() As Long
    If hTabs Then Count = SendMessageW(hTabs, TCM_GETITEMCOUNT, 0, ByVal 0)
End Property

Public Property Get FlatSeparator() As Boolean
    If hTabs Then FlatSeparator = SendMessageW(hTabs, TCM_GETEXTENDEDSTYLE, 0, ByVal 0) And TCS_EX_FLATSEPARATORS
End Property
Public Property Let FlatSeparator(ByVal NewValue As Boolean)
    If hTabs Then Call SendMessageW(hTabs, TCM_SETEXTENDEDSTYLE, 0, ByVal (SendMessageW(hTabs, TCM_GETEXTENDEDSTYLE, 0, ByVal 0) And Not TCS_EX_FLATSEPARATORS) Or (NewValue And TCS_EX_FLATSEPARATORS))
    Call PropertyChanged("Tabs")
End Property

Public Property Get Font() As StdFont
    Set Font = m_oFont
End Property
Public Property Set Font(ByVal New_Font As StdFont)
If hTabs = 0 Then Exit Property
  Dim uLF   As LOGFONT
  Dim nChar As Integer
    Set m_oFont = New_Font
    With uLF
         For nChar = 1 To Len(m_oFont.Name)
             .lfFaceName(nChar - 1) = CByte(Asc(Mid$(m_oFont.Name, nChar, 1)))
         Next nChar
         .lfHeight = -MulDiv(m_oFont.Size, GetDeviceCaps(HDC, LOGPIXELSY), 72)
         .lfItalic = m_oFont.Italic
         If m_oFont.Bold Then .lfWeight = FW_BOLD Else .lfWeight = FW_NORMAL
         .lfUnderline = m_oFont.Underline
         .lfStrikeOut = m_oFont.Strikethrough
         .lfCharSet = m_oFont.Charset
    End With
    Call pvDestroyFont
    m_hFont = CreateFontIndirect(uLF)
    Call PropertyChanged("Font")
    Call SendMessageW(hTabs, WM_SETFONT, m_hFont, ByVal True)
End Property

Public Property Get ItemText(ByVal index As Long) As String
    If hTabs Then
        Dim sTabStrip As TCITEM
        Dim sText As String
        sText = String(255, 0)
        sTabStrip.mask = TCIF_TEXT
        sTabStrip.cchTextMax = 255
        sTabStrip.pszText = StrPtr(sText)
        If SendMessageW(hTabs, TCM_GETITEMW, index, sTabStrip) Then ItemText = Left$(sText, InStr(sText, vbNullChar) - 1)
    End If
End Property
Public Property Let ItemText(ByVal index As Long, ByVal text As String)
    If hTabs Then
        Dim sTabSrip As TCITEM
        sTabSrip.mask = TCIF_TEXT
        sTabSrip.pszText = StrPtr(text)
        Call SendMessageW(hTabs, TCM_SETITEMW, index, sTabSrip)
    End If
End Property

Public Property Get Multiline() As Boolean
    If hTabs Then Multiline = GetWindowLongW(hTabs, GWL_STYLE) And TCS_MULTILINE
End Property
Public Property Let Multiline(ByVal NewValue As Boolean)
    If hTabs Then Call SetWindowLongA(hTabs, GWL_STYLE, (GetWindowLongW(hTabs, GWL_STYLE) And Not TCS_MULTILINE) Or (NewValue And TCS_MULTILINE))
End Property

Public Property Get SelectedItem() As Long
    If hTabs Then SelectedItem = SendMessageW(hTabs, TCM_GETCURSEL, 0, 0)
End Property
Public Property Let SelectedItem(ByVal index As Long)
    Call pvSaveControlState
    If hTabs <> 0 Then Call SendMessageW(hTabs, TCM_SETCURSEL, index, ByVal 0)
    Call pvUpdateTabView
End Property


' === Call InterfaceMethod ===============================================
' This function was made by ANDRay, wich can be found in http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=72856
Private Function CallInterface(ByVal pInterface As Long, ByVal Member As Long, ByVal ParamsCount As Long, Optional ByVal p1 As Long = 0, Optional ByVal p2 As Long = 0, Optional ByVal p3 As Long = 0, Optional ByVal p4 As Long = 0, Optional ByVal p5 As Long = 0, Optional ByVal p6 As Long = 0, Optional ByVal p7 As Long = 0, Optional ByVal p8 As Long = 0, Optional ByVal p9 As Long = 0, Optional ByVal p10 As Long = 0) As Long
  Dim i As Long, t As Long
  Dim hGlobal As Long, hGlobalOffset As Long
  
  If ParamsCount < 0 Then Err.Raise 5 'invalid call
  If pInterface = 0 Then Err.Raise 5
  
  ' 5 Bytes por parametro (4 bytes + PUSH)
  ' 5 Bytes = 1 push + Puntero a interfaz
  hGlobal = GlobalAlloc(GMEM_FIXED, 5 * ParamsCount + 5 + 5 + 3 + 1)
  If hGlobal = 0 Then Err.Raise 7 'insuff. memory
  hGlobalOffset = hGlobal
  
  If ParamsCount > 0 Then
    t = VarPtr(p1)
    For i = ParamsCount - 1 To 0 Step -1
      Call PutMem2(hGlobalOffset, asmPUSH_imm32)
      hGlobalOffset = hGlobalOffset + 1
      Call GetMem4(t + i * 4, hGlobalOffset)
      hGlobalOffset = hGlobalOffset + 4
    Next
  End If
  
  ' PUSH y ponemos el puntero a la interfas
  Call PutMem2(hGlobalOffset, asmPUSH_imm32)
  hGlobalOffset = hGlobalOffset + 1
  Call PutMem4(hGlobalOffset, pInterface)
  hGlobalOffset = hGlobalOffset + 4
  
  ' Llamamos
  Call PutMem2(hGlobalOffset, asmCALL_rel32)
  hGlobalOffset = hGlobalOffset + 1
  Call GetMem4(pInterface, VarPtr(t))     'дереференс: находим положение vTable
  Call GetMem4(t + Member * 4, VarPtr(t)) 'смещение по vTable, после чего дереференс оного
  Call PutMem4(hGlobalOffset, t - hGlobalOffset - 4)
  hGlobalOffset = hGlobalOffset + 4
  
  Call PutMem4(hGlobalOffset, &H10C2&)        'ret 0x0010
  CallInterface = CallWindowProcA(hGlobal, 0, 0, 0, 0)
  Call GlobalFree(hGlobal)
End Function

' === IOLEInPlaceActiveObject Implementation =============================

Private Sub pvInitIPAO()
    Dim uiid As uuid
    ptrMe = ObjPtr(Me)
    With m_uIPAO
        .lpVTable = GetVTable
        Call IIDFromString(StrPtr(IID_IOleInPlaceActive), uiid)
        Call CallInterface(ptrMe, IUnknown_Exports.QueryInterface, 2, VarPtr(uiid), VarPtr(.IPAOReal))
        .ThisPointer = VarPtr(m_uIPAO)
    End With
End Sub

Private Sub pvSetIPAO()
    Const IOleObject_GetClientSite As Long = 4 ' 2 From IUnknown + 2є Ordinal
    Const IOleObject_DoVerb As Long = 11
    Const IOleInPlaceSite_GetWindowContext As Long = 8 ' 2 from IUnknown + 2 IOleWindow + 4є Ordinal
    Const IOleInPlaceFrame_SetActiveObject As Long = 8 ' 2 from IUnknown + 2 IOleWindow + 4є Ordinal
    Const IOleInPlaceUIWindow_SetActiveObject As Long = 8 ' IOleInPlaceFrame inherits from IOleInPlaceUIWindow
    
    Const OLEIVERB_UIACTIVATE As Long = -4
    Dim uiid As uuid, lResult As Long
    Dim pOleObject          As Long 'IOleObject
    Dim pOleInPlaceSite     As Long 'IOleInPlaceSite
    Dim pOleInPlaceFrame    As Long 'IOleInPlaceFrame
    Dim pOleInPlaceUIWindow As Long 'IOleInPlaceUIWindow
    Dim rcPos               As RECT
    Dim rcClip              As RECT
    Dim uFrameInfo          As OLEINPLACEFRAMEINFO
    
    On Error Resume Next
    Call IIDFromString(StrPtr(IID_IOleObject), uiid)
    Call CallInterface(ptrMe, IUnknown_Exports.QueryInterface, 2, VarPtr(uiid), VarPtr(pOleObject))
    Call CallInterface(pOleObject, IOleObject_GetClientSite, 1, VarPtr(pOleInPlaceSite))
    
    If pOleInPlaceSite <> 0 Then
        Call IIDFromString(StrPtr(IID_IOleInPlaceSite), uiid)
        Call CallInterface(pOleInPlaceSite, IUnknown_Exports.QueryInterface, 2, VarPtr(uiid), VarPtr(pOleInPlaceSite))
        Call CallInterface(pOleInPlaceSite, IOleInPlaceSite_GetWindowContext, 5, VarPtr(pOleInPlaceFrame), VarPtr(pOleInPlaceUIWindow), VarPtr(rcPos), VarPtr(rcClip), VarPtr(uFrameInfo))
        
        
        If pOleInPlaceFrame <> 0 Then
            ' The original was pOleInPlaceFrame.SetActiveObject but IOleInPlaceUIWindow has the definition :/
            Call CallInterface(pOleInPlaceFrame, IOleInPlaceFrame_SetActiveObject, 2, m_uIPAO.ThisPointer, StrPtr(vbNullString))
        End If
        If pOleInPlaceUIWindow <> 0 Then  '-- And Not m_bMouseActivate
            Call CallInterface(pOleInPlaceUIWindow, IOleInPlaceUIWindow_SetActiveObject, 2, VarPtr(m_uIPAO.ThisPointer), StrPtr(vbNullString))
        Else
            Call CallInterface(pOleObject, IOleObject_DoVerb, 6, OLEIVERB_UIACTIVATE, 0, pOleInPlaceSite, 0, UserControl.hWnd, VarPtr(rcPos))
        End If
    End If
    
    On Error GoTo 0
End Sub

Private Function pvTranslateAccel(pMsg As msg) As Boolean
    Const IOleObject_GetClientSite As Long = 4 ' 2 From IUnknown + 2є Ordinal
    Dim pOleObject      As Long 'IOleObject
    Dim pOleControlSite As Long 'IOleControlSite
    Dim uiid As uuid
    
    On Error Resume Next
    Select Case pMsg.message
        Case WM_KEYDOWN, WM_KEYUP
            Select Case pMsg.wParam
                Case vbKeyTab
                    If (pvShiftState() And vbCtrlMask) Then
                        Call IIDFromString(StrPtr(IID_IOleObject), uiid)
                        Call CallInterface(ptrMe, IUnknown_Exports.QueryInterface, 2, VarPtr(uiid), VarPtr(pOleObject))
                        Call CallInterface(pOleObject, IOleObject_GetClientSite, 1, VarPtr(pOleControlSite))
                        If pOleControlSite Then
                            Call IIDFromString(StrPtr(IID_IOleControlSite), uiid)
                            Call CallInterface(pOleControlSite, IUnknown_Exports.QueryInterface, 2, VarPtr(uiid), VarPtr(pOleControlSite))
                            Call CallInterface(pOleControlSite, 7, 2, VarPtr(pMsg), pvShiftState() And vbShiftMask)
                        End If
                    End If
                    'frTranslateAccel = False
                Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, vbKeyPageDown, vbKeyPageUp
                    Call SendMessageW(hTabs, pMsg.message, pMsg.wParam, ByVal pMsg.lParam)
                    pvTranslateAccel = True
            End Select
    End Select
    On Error GoTo 0
End Function

Private Function GetVTable() As Long
    ' Set up the vTable for the interface and return a pointer to it
    If (m_IPAOVTable(0) = 0) Then
        m_IPAOVTable(0) = scb_SetCallbackAddr(2, 9)  '9  QueryInterface(2)
        m_IPAOVTable(1) = scb_SetCallbackAddr(1, 11) '11 Addref(1)
        m_IPAOVTable(2) = scb_SetCallbackAddr(1, 10) '10 Release(1)
        m_IPAOVTable(3) = scb_SetCallbackAddr(2, 8)  '8  GetWindow(2)
        m_IPAOVTable(4) = scb_SetCallbackAddr(2, 7)  '7  ContextSensitiveHelp(2)
        m_IPAOVTable(5) = scb_SetCallbackAddr(2, 6)  '6  TranslateAccelerator(2)
        m_IPAOVTable(6) = scb_SetCallbackAddr(2, 5)  '5  OnFrameWindowActivate(2)
        m_IPAOVTable(7) = scb_SetCallbackAddr(2, 4)  '4  OnDocWindowActivate(2)
        m_IPAOVTable(8) = scb_SetCallbackAddr(4, 3)  '3  ResizeBorder(4)
        m_IPAOVTable(9) = scb_SetCallbackAddr(2, 2)  '3  2 EnableModeless(2)
        '--- init guid
        With IID_IOleInPlaceActiveObject
            .Data1 = &H117&
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
    End If
    GetVTable = VarPtr(m_IPAOVTable(0))
End Function


Private Function pvIPAO_AddRef(This As IPAOHookStruct) As Long
    'pvIPAO_AddRef = This.IPAOReal.AddRef
    pvIPAO_AddRef = CallInterface(This.IPAOReal, IUnknown_Exports.AddRef, 0)
End Function

Private Function pvIPAO_Release(This As IPAOHookStruct) As Long
    'pvIPAO_Release = This.IPAOReal.Release
    pvIPAO_Release = CallInterface(This.IPAOReal, IUnknown_Exports.Release, 0)
End Function

Private Function pvIPAO_QueryInterface(This As IPAOHookStruct, riid As uuid, pvObj As Long) As Long
    ' Install the interface if required
    If (IsEqualGUID(riid, IID_IOleInPlaceActiveObject)) Then
        ' Install alternative IOleInPlaceActiveObject interface implemented here
        pvObj = VarPtr(This)
        Call pvIPAO_AddRef(This)
        pvIPAO_QueryInterface = 0
      Else
        ' Use the default support for the interface:
        'pvIPAO_QueryInterface = This.IPAOReal.QueryInterface(ByVal VarPtr(riid), pvObj)
        pvIPAO_QueryInterface = CallInterface(This.IPAOReal, IUnknown_Exports.QueryInterface, 2, VarPtr(riid), VarPtr(pvObj))
    End If
End Function

Private Function pvIPAO_GetWindow(This As IPAOHookStruct, phwnd As Long) As Long
    'pvIPAO_GetWindow = This.IPAOReal.GetWindow(phwnd)
    pvIPAO_GetWindow = CallInterface(This.IPAOReal, IPAO_Exports.GetWindow, 1, VarPtr(phwnd))
End Function

Private Function pvIPAO_ContextSensitiveHelp(This As IPAOHookStruct, ByVal fEnterMode As Long) As Long
    'pvIPAO_ContextSensitiveHelp = This.IPAOReal.ContextSensitiveHelp(fEnterMode)
    pvIPAO_ContextSensitiveHelp = CallInterface(This.IPAOReal, IPAO_Exports.ContextSensitiveHelp, 1, VarPtr(fEnterMode))
End Function

Private Function pvIPAO_TranslateAccelerator(This As IPAOHookStruct, lpMsg As msg) As Long
    ' Check if we want to override the handling of this key code:
    If (pvTranslateAccel(lpMsg)) Then
        pvIPAO_TranslateAccelerator = S_OK
    Else
        ' If not pass it on to the standard UserControl TranslateAccelerator method:
        pvIPAO_TranslateAccelerator = CallInterface(This.IPAOReal, IPAO_Exports.TranslateAccelerator, 1, VarPtr(lpMsg))
        'pvIPAO_TranslateAccelerator = This.IPAOReal.TranslateAccelerator(ByVal VarPtr(lpMsg))
    End If
End Function

Private Function pvIPAO_OnFrameWindowActivate(This As IPAOHookStruct, ByVal fActivate As Long) As Long
    pvIPAO_OnFrameWindowActivate = CallInterface(This.IPAOReal, IPAO_Exports.OnFrameWindowActivate, 1, VarPtr(fActivate))
    'pvIPAO_OnFrameWindowActivate = This.IPAOReal.OnFrameWindowActivate(fActivate)
End Function

Private Function pvIPAO_OnDocWindowActivate(This As IPAOHookStruct, ByVal fActivate As Long) As Long
    pvIPAO_OnDocWindowActivate = CallInterface(This.IPAOReal, IPAO_Exports.OnDocWindowActivate, 1, VarPtr(fActivate))
    'pvIPAO_OnDocWindowActivate = This.IPAOReal.OnDocWindowActivate(fActivate)
End Function

Private Function pvIPAO_ResizeBorder(This As IPAOHookStruct, prcBorder As RECT, ByVal puiWindow As Long, ByVal fFrameWindow As Long) As Long
    'pvIPAO_ResizeBorder = This.IPAOReal.ResizeBorder(VarPtr(prcBorder), puiWindow, fFrameWindow)
    pvIPAO_ResizeBorder = CallInterface(This.IPAOReal, IPAO_Exports.ResizeBorder, 3, VarPtr(prcBorder), puiWindow, VarPtr(fFrameWindow))
End Function

Private Function pvIPAO_EnableModeless(This As IPAOHookStruct, ByVal fEnable As Long) As Long
    'pvIPAO_EnableModeless = This.IPAOReal.EnableModeless(fEnable)
    pvIPAO_EnableModeless = CallInterface(This.IPAOReal, IPAO_Exports.EnableModeless, 1, VarPtr(fEnable))
End Function
' === Subclassing callback ===============================================
Private Sub WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, _
                      ByRef lParamUser As Long)
    Dim snx      As Single
    Dim sny      As Single
    Select Case lng_hWnd
        Case UserControl.hWnd
            Select Case uMsg
                Case WM_NOTIFY
                    Dim uNMH        As NMHDR
                    Call RtlMoveMemory(VarPtr(uNMH), lParam, Len(uNMH))
                    Select Case uNMH.code
                        Case NM_CLICK
                            Call pvUCCoordScale(snx, sny)
                            RaiseEvent MouseUp((uNMH.code = NM_CLICK) + 2, pvShiftState(), snx, sny)
                            RaiseEvent Click
                        Case NM_DBLCLK, NM_RDBLCLK
                            RaiseEvent DblClick
                        Case NM_RCLICK
                            Dim ht As TCHITTESTINFO
                            Call GetCursorPos(ht.pt)
                            Call ScreenToClient(hTabs, ht.pt)
                            ht.flags = TCHT_ONITEM
                            RaiseEvent TabRightClick(SendMessageW(hTabs, TCM_HITTEST, 0, ht))
                        Case TCN_SELCHANGE
                            RaiseEvent TabClick(SelectedItem)
                            ' Changin items! :D
                            Call pvUpdateTabView
                        Case TCN_SELCHANGING
                            Dim b As Boolean
                            RaiseEvent ChangingTab(b)
                            lReturn = b
                            bHandled = True
                    End Select
                Case WM_MOUSEACTIVATE
                    'Call pvSetIPAO
            End Select
            'WndProc = CallWindowProcW(hOldWndProc, lng_hWnd, uMsg, wParam, lParam)
        Case hTabs
            Select Case uMsg
                Case WM_SETFOCUS
                   Call pvSetIPAO
                Case WM_DESTROY
                    hTabs = 0
                Case WM_KEYDOWN
                   RaiseEvent KeyDown(wParam And &H7FFF&, pvShiftState())
               Case WM_CHAR
                   RaiseEvent KeyPress(wParam And &H7FFF&)
               Case WM_KEYUP
                   RaiseEvent KeyUp(wParam And &H7FFF&, pvShiftState())
               Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN
                   Call pvUCCoordScale(snx, sny)
                   RaiseEvent MouseDown(pvButton(uMsg), pvShiftState(), snx, sny)
               Case WM_MOUSEMOVE
                   If (Not m_bInCtrl) Then
                       m_bInCtrl = True
                       Call pvTrackMouseLeave(lng_hWnd)
                       RaiseEvent MouseEnter
                   End If
                   Call pvUCCoordScale(snx, sny)
                   If (snx <> m_snxL Or sny <> m_snyL) Then
                       m_snxL = snx
                       m_snyL = sny
                       RaiseEvent MouseMove(pvButton(uMsg), pvShiftState(), snx, sny)
                   End If
               Case WM_MOUSELEAVE
                   m_bInCtrl = False
                   RaiseEvent MouseLeave
                   m_snxL = -1
                   m_snyL = -1
               Case WM_LBUTTONUP, WM_RBUTTONUP, WM_MBUTTONUP
                   Call pvUCCoordScale(snx, sny)
                   RaiseEvent MouseUp(pvButton(uMsg), pvShiftState(), snx, sny)
                   'RaiseEvent Click
                Case TCM_DELETEALLITEMS
                    SelectedItem = -1
            End Select
            'WndProc = CallWindowProcW(hOldTabWndProc, lng_hWnd, uMsg, wParam, lParam)
    End Select
End Sub
