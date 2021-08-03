Option Strict Off
Option Explicit On
Module VisualStyles
	Public Declare Function ActivateVisualStyles Lib "uxtheme"  Alias "SetWindowTheme"(ByVal hWnd As Integer, Optional ByVal pszSubAppName As Integer = 0, Optional ByVal pszSubIdList As Integer = 0) As Integer
	Public Declare Function RemoveVisualStyles Lib "uxtheme"  Alias "SetWindowTheme"(ByVal hWnd As Integer, Optional ByRef pszSubAppName As String = " ", Optional ByRef pszSubIdList As String = " ") As Integer
	Public Declare Function GetVisualStyles Lib "uxtheme"  Alias "GetWindowTheme"(ByVal hWnd As Integer) As Integer
	Private Structure TINITCOMMONCONTROLSEX
		Dim dwSize As Integer
		Dim dwICC As Integer
	End Structure
	Private Structure TRELEASE
		Dim IUnk As stdole.IUnknown
		<VBFixedArray(2)> Dim VTable() As Integer
		Dim VTableHeaderPointer As Integer
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim VTable(2)
		End Sub
	End Structure
	Private Structure TRACKMOUSEEVENTSTRUCT
		Dim cbSize As Integer
		Dim dwFlags As Integer
		Dim hWndTrack As Integer
		Dim dwHoverTime As Integer
	End Structure
	Private Enum UxThemeButtonParts
		BP_PUSHBUTTON = 1
		BP_RADIOBUTTON = 2
		BP_CHECKBOX = 3
		BP_GROUPBOX = 4
		BP_USERBUTTON = 5
	End Enum
	Private Enum UxThemeButtonStates
		PBS_NORMAL = 1
		PBS_HOT = 2
		PBS_PRESSED = 3
		PBS_DISABLED = 4
		PBS_DEFAULTED = 5
	End Enum
	Private Structure POINTAPI
		Dim X As Integer
		Dim Y As Integer
	End Structure
	Private Structure RECT
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
	Private Structure PAINTSTRUCT
		Dim hDC As Integer
		Dim fErase As Integer
		Dim RCPaint As RECT
		Dim fRestore As Integer
		Dim fIncUpdate As Integer
		<VBFixedArray(31)> Dim RGBReserved() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim RGBReserved(31)
		End Sub
	End Structure
	Private Structure DLLVERSIONINFO
		Dim cbSize As Integer
		Dim dwMajor As Integer
		Dim dwMinor As Integer
		Dim dwBuildNumber As Integer
		Dim dwPlatformID As Integer
	End Structure
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef Destination As Any, ByRef Source As Any, ByVal Length As Integer)
	'UPGRADE_WARNING: Structure TINITCOMMONCONTROLSEX may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function InitCommonControlsEx Lib "comctl32" (ByRef ICCEX As TINITCOMMONCONTROLSEX) As Integer
	Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Integer) As Integer
	'UPGRADE_WARNING: Structure DLLVERSIONINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DllGetVersion Lib "comctl32" (ByRef pdvi As DLLVERSIONINFO) As Integer
	Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Integer) As Integer
	Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Integer) As Integer
	Private Declare Function GetFocus Lib "user32" () As Integer
	'UPGRADE_WARNING: Structure POINTAPI may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Integer
	Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Integer, ByVal Y As Integer) As Integer
	Private Declare Function ExtSelectClipRgn Lib "gdi32" (ByVal hDC As Integer, ByVal hRgn As Integer, ByVal fnMode As Integer) As Integer
	Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
	Private Declare Function DrawState Lib "user32"  Alias "DrawStateW"(ByVal hDC As Integer, ByVal hBrush As Integer, ByVal lpDrawStateProc As Integer, ByVal lData As Integer, ByVal wData As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal CX As Integer, ByVal CY As Integer, ByVal fFlags As Integer) As Integer
	Private Declare Function GetProp Lib "user32"  Alias "GetPropW"(ByVal hWnd As Integer, ByVal lpString As Integer) As Integer
	Private Declare Function SetProp Lib "user32"  Alias "SetPropW"(ByVal hWnd As Integer, ByVal lpString As Integer, ByVal hData As Integer) As Integer
	Private Declare Function RemoveProp Lib "user32"  Alias "RemovePropW"(ByVal hWnd As Integer, ByVal lpString As Integer) As Integer
	'UPGRADE_WARNING: Structure PAINTSTRUCT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Integer, ByRef lpPaint As PAINTSTRUCT) As Integer
	'UPGRADE_WARNING: Structure PAINTSTRUCT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Integer, ByRef lpPaint As PAINTSTRUCT) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Integer, ByRef lpRect As RECT) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Integer, ByRef lpRect As Any, ByVal bErase As Integer) As Integer
	Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function FillRect Lib "user32" (ByVal hDC As Integer, ByRef lpRect As RECT, ByVal hBrush As Integer) As Integer
	Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Integer) As Integer
	Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Integer, ByVal hDC As Integer) As Integer
	Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Integer) As Integer
	Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Integer) As Integer
	Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Integer, ByRef lpRect As RECT) As Integer
	Private Declare Function GetDC Lib "user32" (ByVal hWnd As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DrawText Lib "user32"  Alias "DrawTextW"(ByVal hDC As Integer, ByVal lpStr As Integer, ByVal nCount As Integer, ByRef lpRect As RECT, ByVal wFormat As Integer) As Integer
	'UPGRADE_WARNING: Structure TRACKMOUSEEVENTSTRUCT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENTSTRUCT) As Integer
	Private Declare Function TransparentBlt Lib "msimg32" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal nWidthSrc As Integer, ByVal nHeightSrc As Integer, ByVal crTransparent As Integer) As Integer
	Private Declare Function IsThemeBackgroundPartiallyTransparent Lib "uxtheme" (ByVal Theme As Integer, ByRef iPartId As Integer, ByRef iStateId As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DrawThemeParentBackground Lib "uxtheme" (ByVal hWnd As Integer, ByVal hDC As Integer, ByRef pRect As RECT) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal Theme As Integer, ByVal hDC As Integer, ByVal iPartId As Integer, ByVal iStateId As Integer, ByRef pRect As RECT, ByRef pClipRect As RECT) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function DrawThemeText Lib "uxtheme" (ByVal Theme As Integer, ByVal hDC As Integer, ByVal iPartId As Integer, ByVal iStateId As Integer, ByVal pszText As Integer, ByVal iCharCount As Integer, ByVal dwTextFlags As Integer, ByVal dwTextFlags2 As Integer, ByRef pRect As RECT) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetThemeBackgroundRegion Lib "uxtheme" (ByVal Theme As Integer, ByVal hDC As Integer, ByVal iPartId As Integer, ByVal iStateId As Integer, ByRef pRect As RECT, ByRef hRgn As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme" (ByVal Theme As Integer, ByVal hDC As Integer, ByVal iPartId As Integer, ByVal iStateId As Integer, ByRef pBoundingRect As RECT, ByRef pContentRect As RECT) As Integer
	Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Integer, ByVal pszClassList As Integer) As Integer
	Private Declare Function CloseThemeData Lib "uxtheme" (ByVal Theme As Integer) As Integer
	Private Declare Function IsAppThemed Lib "uxtheme" () As Integer
	Private Declare Function IsThemeActive Lib "uxtheme" () As Integer
	Private Declare Function GetThemeAppProperties Lib "uxtheme" () As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function SendMessage Lib "user32"  Alias "SendMessageW"(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Any) As Integer
	Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Integer, ByVal pfnSubclass As Integer, ByVal uIdSubclass As Integer, ByVal dwRefData As Integer) As Integer
	Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Integer, ByVal pfnSubclass As Integer, ByVal uIdSubclass As Integer) As Integer
	Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
	Private Declare Function DefWindowProc Lib "user32"  Alias "DefWindowProcW"(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
	Private Const ICC_STANDARD_CLASSES As Integer = &H4000
	Private Const STAP_ALLOW_CONTROLS As Integer = (1 * (2 ^ 1))
	Private Const S_OK As Integer = &H0
	Private Const UIS_CLEAR As Integer = 2
	Private Const UISF_HIDEFOCUS As Integer = &H1
	Private Const UISF_HIDEACCEL As Integer = &H2
	Private Const WM_UPDATEUISTATE As Integer = &H128
	Private Const WM_QUERYUISTATE As Integer = &H129
	Private Const WM_SETFOCUS As Integer = &H7
	Private Const WM_KILLFOCUS As Integer = &H8
	Private Const WM_ENABLE As Integer = &HA
	Private Const WM_SETREDRAW As Integer = &HB
	Private Const WM_PAINT As Integer = &HF
	Private Const WM_NCPAINT As Integer = &H85
	Private Const WM_NCDESTROY As Integer = &H82
	Private Const BM_GETSTATE As Integer = &HF2
	Private Const WM_MOUSEMOVE As Integer = &H200
	Private Const WM_LBUTTONDOWN As Integer = &H201
	Private Const WM_LBUTTONUP As Integer = &H202
	Private Const WM_RBUTTONUP As Integer = &H205
	Private Const WM_MOUSELEAVE As Integer = &H2A3
	Private Const WM_PRINTCLIENT As Integer = &H318
	Private Const WM_THEMECHANGED As Integer = &H31A
	Private Const BST_PUSHED As Integer = &H4
	Private Const BST_FOCUS As Integer = &H8
	Private Const DT_CENTER As Integer = &H1
	Private Const DT_WORDBREAK As Integer = &H10
	Private Const DT_CALCRECT As Integer = &H400
	Private Const DT_HIDEPREFIX As Integer = &H100000
	Private Const TME_LEAVE As Integer = 2
	Private Const RGN_DIFF As Integer = 4
	Private Const RGN_COPY As Integer = 5
	Private Const DST_ICON As Integer = &H3
	Private Const DST_BITMAP As Integer = &H4
	Private Const DSS_DISABLED As Integer = &H20
	
	Public Sub InitVisualStylesFixes()
		'UPGRADE_ISSUE: App property App.LogMode was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		'UPGRADE_WARNING: Add a delegate for AddressOf ReleaseVisualStylesFixes Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
		If App.LogMode <> 0 Then Call InitReleaseVisualStylesFixes(AddressOf ReleaseVisualStylesFixes)
		Dim ICCEX As TINITCOMMONCONTROLSEX
		With ICCEX
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			.dwSize = LenB(ICCEX)
			.dwICC = ICC_STANDARD_CLASSES
		End With
		InitCommonControlsEx(ICCEX)
	End Sub
	
	Private Sub InitReleaseVisualStylesFixes(ByVal Address As Integer)
		'UPGRADE_WARNING: Arrays in structure Release may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Static Release As TRELEASE
		If Release.VTableHeaderPointer <> 0 Then Exit Sub
		If GetComCtlVersion >= 6 Then
			Release.VTable(2) = Address
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			Release.VTableHeaderPointer = VarPtr(Release.VTable(0))
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Release.IUnk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(Release.IUnk, VarPtr(Release.VTableHeaderPointer), 4)
		End If
	End Sub
	
	Private Function ReleaseVisualStylesFixes() As Integer
		Const SEM_NOGPFAULTERRORBOX As Integer = &H2
		SetErrorMode(SEM_NOGPFAULTERRORBOX)
	End Function
	
	Public Sub SetupVisualStylesFixes(ByVal Form As System.Windows.Forms.Form)
		If GetComCtlVersion() >= 6 Then SendMessage(Form.Handle.ToInt32, WM_UPDATEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEFOCUS Or UISF_HIDEACCEL), 0)
		If EnabledVisualStyles() = False Then Exit Sub
		Dim CurrControl As System.Windows.Forms.Control
		For	Each CurrControl In Form.Controls
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			Select Case TypeName(CurrControl)
				Case "Frame"
					'UPGRADE_ISSUE: ObjPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					'UPGRADE_WARNING: Add a delegate for AddressOf RedirectFrame Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
					SetWindowSubclass(CurrControl.Handle.ToInt32, AddressOf RedirectFrame, ObjPtr(CurrControl), 0)
				Case "CommandButton", "OptionButton", "CheckBox"
					'UPGRADE_WARNING: Couldn't resolve default property of object CurrControl.Style. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CurrControl.Style = System.Windows.Forms.Appearance.Button Then
						'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
						SetProp(CurrControl.Handle.ToInt32, StrPtr("VisualStyles"), GetVisualStyles(CurrControl.Handle.ToInt32))
						'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
						If CurrControl.Enabled = True Then SetProp(CurrControl.Handle.ToInt32, StrPtr("Enabled"), 1)
						'UPGRADE_ISSUE: ObjPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
						'UPGRADE_WARNING: Add a delegate for AddressOf RedirectButton Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
						SetWindowSubclass(CurrControl.Handle.ToInt32, AddressOf RedirectButton, ObjPtr(CurrControl), ObjPtr(CurrControl))
					End If
			End Select
		Next CurrControl
	End Sub
	
	Public Sub RemoveVisualStylesFixes(ByVal Form As System.Windows.Forms.Form)
		If EnabledVisualStyles() = False Then Exit Sub
		Dim CurrControl As System.Windows.Forms.Control
		For	Each CurrControl In Form.Controls
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			Select Case TypeName(CurrControl)
				Case "Frame"
					'UPGRADE_ISSUE: ObjPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					'UPGRADE_WARNING: Add a delegate for AddressOf RedirectFrame Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
					RemoveWindowSubclass(CurrControl.Handle.ToInt32, AddressOf RedirectFrame, ObjPtr(CurrControl))
				Case "CommandButton", "OptionButton", "CheckBox"
					'UPGRADE_WARNING: Couldn't resolve default property of object CurrControl.Style. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If CurrControl.Style = System.Windows.Forms.Appearance.Button Then
						'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
						RemoveProp(CurrControl.Handle.ToInt32, StrPtr("VisualStyles"))
						'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
						RemoveProp(CurrControl.Handle.ToInt32, StrPtr("Enabled"))
						'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
						RemoveProp(CurrControl.Handle.ToInt32, StrPtr("Hot"))
						'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
						RemoveProp(CurrControl.Handle.ToInt32, StrPtr("Painted"))
						'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
						RemoveProp(CurrControl.Handle.ToInt32, StrPtr("ButtonPart"))
						'UPGRADE_ISSUE: ObjPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
						'UPGRADE_WARNING: Add a delegate for AddressOf RedirectButton Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
						RemoveWindowSubclass(CurrControl.Handle.ToInt32, AddressOf RedirectButton, ObjPtr(CurrControl))
					End If
			End Select
		Next CurrControl
	End Sub
	
	Public Function EnabledVisualStyles() As Boolean
		If GetComCtlVersion() >= 6 Then
			If IsThemeActive() <> 0 Then
				If IsAppThemed() <> 0 Then
					EnabledVisualStyles = True
				ElseIf (GetThemeAppProperties() And STAP_ALLOW_CONTROLS) <> 0 Then 
					EnabledVisualStyles = True
				End If
			End If
		End If
	End Function
	
	Public Function GetComCtlVersion() As Integer
		Static Done As Boolean
		Static Value As Integer
		Dim Version As DLLVERSIONINFO
		If Done = False Then
			On Error Resume Next
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			Version.cbSize = LenB(Version)
			If DllGetVersion(Version) = S_OK Then Value = Version.dwMajor
			Done = True
		End If
		GetComCtlVersion = Value
	End Function
	
	Private Function RedirectFrame(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer, ByVal uIdSubclass As Integer, ByVal dwRefData As Integer) As Integer
		Select Case wMsg
			Case WM_PRINTCLIENT, WM_MOUSELEAVE
				RedirectFrame = DefWindowProc(hWnd, wMsg, wParam, lParam)
				Exit Function
		End Select
		RedirectFrame = DefSubclassProc(hWnd, wMsg, wParam, lParam)
		If wMsg = WM_NCDESTROY Then Call RemoveRedirectFrame(hWnd, uIdSubclass)
	End Function
	
	Private Sub RemoveRedirectFrame(ByVal hWnd As Integer, ByVal uIdSubclass As Integer)
		'UPGRADE_WARNING: Add a delegate for AddressOf RedirectFrame Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
		RemoveWindowSubclass(hWnd, AddressOf RedirectFrame, uIdSubclass)
	End Sub
	
	Private Function RedirectButton(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer, ByVal uIdSubclass As Integer, ByVal Button As Object) As Integer
		Dim SetRedraw As Boolean
		'UPGRADE_WARNING: Arrays in structure PS may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim PS As PAINTSTRUCT
		Select Case wMsg
			Case WM_NCPAINT
				Exit Function
			Case WM_PAINT
				'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				If IsWindowVisible(hWnd) <> 0 And GetProp(hWnd, StrPtr("VisualStyles")) <> 0 Then
					'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					SetProp(hWnd, StrPtr("Painted"), 1)
					Call DrawButton(hWnd, BeginPaint(hWnd, PS), Button)
					EndPaint(hWnd, PS)
					Exit Function
				End If
			Case WM_SETFOCUS, WM_ENABLE
				If IsWindowVisible(hWnd) <> 0 Then
					SetRedraw = True
					SendMessage(hWnd, WM_SETREDRAW, 0, 0)
				End If
		End Select
		RedirectButton = DefSubclassProc(hWnd, wMsg, wParam, lParam)
		Dim TME As TRACKMOUSEEVENTSTRUCT
		If wMsg = WM_NCDESTROY Then
			Call RemoveRedirectButton(hWnd, uIdSubclass)
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			RemoveProp(hWnd, StrPtr("VisualStyles"))
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			RemoveProp(hWnd, StrPtr("Enabled"))
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			RemoveProp(hWnd, StrPtr("Hot"))
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			RemoveProp(hWnd, StrPtr("Painted"))
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			RemoveProp(hWnd, StrPtr("ButtonPart"))
		ElseIf IsWindow(hWnd) <> 0 Then 
			Select Case wMsg
				Case WM_THEMECHANGED
					'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					SetProp(hWnd, StrPtr("VisualStyles"), GetVisualStyles(hWnd))
					'UPGRADE_WARNING: Couldn't resolve default property of object Button.Refresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Button.Refresh()
				Case WM_MOUSELEAVE
					'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					SetProp(hWnd, StrPtr("Hot"), 0)
					'UPGRADE_WARNING: Couldn't resolve default property of object Button.Refresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Button.Refresh()
				Case WM_MOUSEMOVE
					'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					If GetProp(hWnd, StrPtr("Hot")) = 0 Then
						'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
						SetProp(hWnd, StrPtr("Hot"), 1)
						InvalidateRect(hWnd, 0, 0)
						With TME
							'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
							.cbSize = LenB(TME)
							.hWndTrack = hWnd
							.dwFlags = TME_LEAVE
						End With
						TrackMouseEvent(TME)
						'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					ElseIf GetProp(hWnd, StrPtr("Painted")) = 0 Then 
						'UPGRADE_WARNING: Couldn't resolve default property of object Button.Refresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Button.Refresh()
					End If
				Case WM_SETFOCUS, WM_ENABLE
					If SetRedraw = True Then
						SendMessage(hWnd, WM_SETREDRAW, 1, 0)
						If wMsg = WM_ENABLE Then
							'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
							SetProp(hWnd, StrPtr("Enabled"), 0)
							InvalidateRect(hWnd, 0, 0)
						Else
							'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
							SetProp(hWnd, StrPtr("Enabled"), 1)
							'UPGRADE_WARNING: Couldn't resolve default property of object Button.Refresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							Button.Refresh()
						End If
					End If
				Case WM_LBUTTONDOWN, WM_LBUTTONUP, WM_RBUTTONUP
					'UPGRADE_WARNING: Couldn't resolve default property of object Button.Refresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Button.Refresh()
			End Select
		End If
	End Function
	
	Private Sub RemoveRedirectButton(ByVal hWnd As Integer, ByVal uIdSubclass As Integer)
		'UPGRADE_WARNING: Add a delegate for AddressOf RedirectButton Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
		RemoveWindowSubclass(hWnd, AddressOf RedirectButton, uIdSubclass)
	End Sub
	
	Private Sub DrawButton(ByVal hWnd As Integer, ByVal hDC As Integer, ByVal Button As Object)
		Dim ButtonState, Theme, ButtonPart, UIState As Integer
		'UPGRADE_NOTE: Default was upgraded to Default_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Pushed, Default_Renamed, Enabled, Checked, Hot, Focused As Boolean
		Dim hFontOld As Integer
		Dim ButtonFont As System.Drawing.Font
		Dim ButtonPicture As System.Drawing.Image
		Dim DisabledPictureAvailable As Boolean
		Dim ClientRect, TextRect As RECT
		Dim RgnClip As Integer
		Dim X, CX, CY, Y As Integer
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		ButtonPart = GetProp(hWnd, StrPtr("ButtonPart"))
		If ButtonPart = 0 Then
			'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			Select Case TypeName(Button)
				Case "CommandButton"
					ButtonPart = UxThemeButtonParts.BP_PUSHBUTTON
				Case "OptionButton"
					ButtonPart = UxThemeButtonParts.BP_RADIOBUTTON
				Case "CheckBox"
					ButtonPart = UxThemeButtonParts.BP_CHECKBOX
			End Select
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If ButtonPart <> 0 Then SetProp(hWnd, StrPtr("ButtonPart"), ButtonPart)
		End If
		Select Case ButtonPart
			Case UxThemeButtonParts.BP_PUSHBUTTON
				'UPGRADE_WARNING: Couldn't resolve default property of object Button.Default. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Default_Renamed = Button.Default
				If GetFocus() <> hWnd Then
					On Error Resume Next
					'UPGRADE_WARNING: Couldn't resolve default property of object Button.Parent. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_ISSUE: Control Default could not be resolved because it was within the generic namespace ActiveControl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
					If CInt(Button.Parent.ActiveControl.Default) > 0 Then 
					Else 
						Default_Renamed = False
					End If
					On Error GoTo 0
				End If
			Case UxThemeButtonParts.BP_RADIOBUTTON
				'UPGRADE_WARNING: Couldn't resolve default property of object Button.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Checked = Button.Value
				Default_Renamed = False
			Case UxThemeButtonParts.BP_CHECKBOX
				'UPGRADE_WARNING: Couldn't resolve default property of object Button.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Checked = IIf(Button.Value = System.Windows.Forms.CheckState.Checked, True, False)
				Default_Renamed = False
		End Select
		ButtonPart = UxThemeButtonParts.BP_PUSHBUTTON
		ButtonState = SendMessage(hWnd, BM_GETSTATE, 0, 0)
		UIState = SendMessage(hWnd, WM_QUERYUISTATE, 0, 0)
		'UPGRADE_WARNING: Couldn't resolve default property of object Button.Enabled. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		Enabled = IIf(GetProp(hWnd, StrPtr("Enabled")) = 1, True, Button.Enabled)
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		Hot = IIf(GetProp(hWnd, StrPtr("Hot")) = 0, False, True)
		If Checked = True Then Hot = False
		Pushed = IIf((ButtonState And BST_PUSHED) = 0, False, True)
		Focused = IIf((ButtonState And BST_FOCUS) = 0, False, True)
		If Enabled = False Then
			ButtonState = UxThemeButtonStates.PBS_DISABLED
			'UPGRADE_WARNING: Couldn't resolve default property of object Button.Picture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Button.DisabledPicture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ButtonPicture = CoalescePicture(Button.DisabledPicture, Button.Picture)
			'UPGRADE_WARNING: Couldn't resolve default property of object Button.DisabledPicture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not Button.DisabledPicture Is Nothing Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Button.DisabledPicture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Button.DisabledPicture.Handle <> 0 Then DisabledPictureAvailable = True
			End If
		ElseIf Hot = True And Pushed = False Then 
			ButtonState = UxThemeButtonStates.PBS_HOT
			If Checked = True Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Button.Picture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object Button.DownPicture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ButtonPicture = CoalescePicture(Button.DownPicture, Button.Picture)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Button.Picture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ButtonPicture = Button.Picture
			End If
		ElseIf Checked = True Or Pushed = True Then 
			ButtonState = UxThemeButtonStates.PBS_PRESSED
			'UPGRADE_WARNING: Couldn't resolve default property of object Button.Picture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object Button.DownPicture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ButtonPicture = CoalescePicture(Button.DownPicture, Button.Picture)
		ElseIf Focused = True Or Default_Renamed = True Then 
			ButtonState = UxThemeButtonStates.PBS_DEFAULTED
			'UPGRADE_WARNING: Couldn't resolve default property of object Button.Picture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ButtonPicture = Button.Picture
		Else
			ButtonState = UxThemeButtonStates.PBS_NORMAL
			'UPGRADE_WARNING: Couldn't resolve default property of object Button.Picture. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ButtonPicture = Button.Picture
		End If
		If Not ButtonPicture Is Nothing Then
			'UPGRADE_ISSUE: Picture property ButtonPicture.Handle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_NOTE: Object ButtonPicture may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			If ButtonPicture.Handle = 0 Then ButtonPicture = Nothing
		End If
		GetClientRect(hWnd, ClientRect)
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		Theme = OpenThemeData(hWnd, StrPtr("Button"))
		Dim Brush As Integer
		If Theme <> 0 Then
			GetThemeBackgroundRegion(Theme, hDC, ButtonPart, ButtonState, ClientRect, RgnClip)
			ExtSelectClipRgn(hDC, RgnClip, RGN_DIFF)
			'UPGRADE_WARNING: Couldn't resolve default property of object Button.BackColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Brush = CreateSolidBrush(WinColor(Button.BackColor))
			FillRect(hDC, ClientRect, Brush)
			DeleteObject(Brush)
			If IsThemeBackgroundPartiallyTransparent(Theme, ButtonPart, ButtonState) <> 0 Then DrawThemeParentBackground(hWnd, hDC, ClientRect)
			ExtSelectClipRgn(hDC, 0, RGN_COPY)
			DeleteObject(RgnClip)
			DrawThemeBackground(Theme, hDC, ButtonPart, ButtonState, ClientRect, ClientRect)
			GetThemeBackgroundContentRect(Theme, hDC, ButtonPart, ButtonState, ClientRect, ClientRect)
			If Focused = True Then
				If Not (UIState And UISF_HIDEFOCUS) = UISF_HIDEFOCUS Then DrawFocusRect(hDC, ClientRect)
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object Button.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not Button.Caption = vbNullString Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Button.Font. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ButtonFont = Button.Font
				hFontOld = SelectObject(hDC, ButtonFont.hFont)
				'UPGRADE_ISSUE: LSet cannot assign one type to another. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="899FA812-8F71-4014-BAEE-E5AF348BA5AA"'
				TextRect = LSet(ClientRect)
				'UPGRADE_WARNING: Couldn't resolve default property of object Button.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				DrawText(hDC, StrPtr(Button.Caption), -1, TextRect, DT_CALCRECT Or DT_WORDBREAK Or CInt(IIf((UIState And UISF_HIDEACCEL) = UISF_HIDEACCEL, DT_HIDEPREFIX, 0)))
				TextRect.Left_Renamed = ClientRect.Left_Renamed
				TextRect.Right_Renamed = ClientRect.Right_Renamed
				If ButtonPicture Is Nothing Then
					TextRect.Top = ((ClientRect.Bottom - TextRect.Bottom) / 2) + (3 * PixelsPerDIP_Y())
					TextRect.Bottom = TextRect.Top + TextRect.Bottom
				Else
					TextRect.Top = (ClientRect.Bottom - TextRect.Bottom) + (1 * PixelsPerDIP_Y())
					TextRect.Bottom = ClientRect.Bottom
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object Button.Caption. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				DrawThemeText(Theme, hDC, ButtonPart, ButtonState, StrPtr(Button.Caption), -1, DT_CENTER Or DT_WORDBREAK Or CInt(IIf((UIState And UISF_HIDEACCEL) = UISF_HIDEACCEL, DT_HIDEPREFIX, 0)), 0, TextRect)
				SelectObject(hDC, hFontOld)
				ClientRect.Bottom = TextRect.Top
				ClientRect.Left_Renamed = TextRect.Left_Renamed
			End If
			CloseThemeData(Theme)
		End If
		Dim hDCScreen As Integer
		Dim hDC1, hBmpOld1 As Integer
		Dim hImage As Integer
		If Not ButtonPicture Is Nothing Then
			'UPGRADE_ISSUE: Picture property ButtonPicture.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			CX = CHimetricToPixel_X(ButtonPicture.Width)
			'UPGRADE_ISSUE: Picture property ButtonPicture.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			CY = CHimetricToPixel_Y(ButtonPicture.Height)
			X = ClientRect.Left_Renamed + ((ClientRect.Right_Renamed - ClientRect.Left_Renamed - CX) / 2)
			Y = ClientRect.Top + ((ClientRect.Bottom - ClientRect.Top - CY) / 2)
			If Enabled = True Or DisabledPictureAvailable = True Then
				'UPGRADE_WARNING: Couldn't resolve default property of object Button.UseMaskColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_ISSUE: Constant vbPicTypeBitmap was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
				'UPGRADE_ISSUE: Picture property ButtonPicture.Type was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				If ButtonPicture.Type = vbPicTypeBitmap And Button.UseMaskColor = True Then
					hDCScreen = GetDC(0)
					If hDCScreen <> 0 Then
						hDC1 = CreateCompatibleDC(hDCScreen)
						If hDC1 <> 0 Then
							'UPGRADE_ISSUE: Picture property ButtonPicture.Handle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							hBmpOld1 = SelectObject(hDC1, ButtonPicture.Handle)
							'UPGRADE_WARNING: Couldn't resolve default property of object Button.MaskColor. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							TransparentBlt(hDC, X, Y, CX, CY, hDC1, 0, 0, CX, CY, WinColor(Button.MaskColor))
							SelectObject(hDC1, hBmpOld1)
							DeleteDC(hDC1)
						End If
						ReleaseDC(0, hDCScreen)
					End If
				Else
					With ButtonPicture
						'UPGRADE_ISSUE: Picture property ButtonPicture.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_ISSUE: Picture property ButtonPicture.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_ISSUE: Picture method ButtonPicture.Render was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						.Render(hDC Or 0, X Or 0, Y Or 0, CX Or 0, CY Or 0, 0, .Height, .Width, -.Height, 0)
					End With
				End If
			Else
				'UPGRADE_ISSUE: Constant vbPicTypeIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
				'UPGRADE_ISSUE: Picture property ButtonPicture.Type was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				If ButtonPicture.Type = vbPicTypeIcon Then
					'UPGRADE_ISSUE: Picture property ButtonPicture.Handle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					DrawState(hDC, 0, 0, ButtonPicture.Handle, 0, X, Y, CX, CY, DST_ICON Or DSS_DISABLED)
				Else
					hImage = BitmapHandleFromPicture(ButtonPicture, System.Drawing.Color.White)
					' The DrawState API with DSS_DISABLED will draw white as transparent.
					' This will ensure GIF bitmaps or metafiles are better drawn.
					DrawState(hDC, 0, 0, hImage, 0, X, Y, CX, CY, DST_BITMAP Or DSS_DISABLED)
					DeleteObject(hImage)
				End If
			End If
		End If
	End Sub
	
	Private Function CoalescePicture(ByVal Picture As System.Drawing.Image, ByVal DefaultPicture As System.Drawing.Image) As System.Drawing.Image
		If Picture Is Nothing Then
			CoalescePicture = DefaultPicture
			'UPGRADE_ISSUE: Picture property Picture.Handle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		ElseIf Picture.Handle = 0 Then 
			CoalescePicture = DefaultPicture
		Else
			CoalescePicture = Picture
		End If
	End Function
End Module