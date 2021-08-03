Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module Common
	Private Structure MSGBOXPARAMS
		Dim cbSize As Integer
		Dim hWndOwner As Integer
		Dim hInstance As Integer
		Dim lpszText As Integer
		Dim lpszCaption As Integer
		Dim dwStyle As Integer
		Dim lpszIcon As Integer
		Dim dwContextHelpID As Integer
		Dim lpfnMsgBoxCallback As Integer
		Dim dwLanguageId As Integer
	End Structure
	Private Structure RECT
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
	Private Structure POINTAPI
		Dim X As Integer
		Dim Y As Integer
	End Structure
	Private Structure BITMAP
		Dim BMType As Integer
		Dim BMWidth As Integer
		Dim BMHeight As Integer
		Dim BMWidthBytes As Integer
		Dim BMPlanes As Short
		Dim BMBitsPixel As Short
		Dim BMBits As Integer
	End Structure
	Private Structure SAFEARRAYBOUND
		Dim cElements As Integer
		Dim lLbound As Integer
	End Structure
	Private Structure SAFEARRAY1D
		Dim cDims As Short
		Dim fFeatures As Short
		Dim cbElements As Integer
		Dim cLocks As Integer
		Dim pvData As Integer
		Dim Bounds As SAFEARRAYBOUND
	End Structure
	Private Structure PICTDESC
		Dim cbSizeOfStruct As Integer
		Dim PicType As Integer
		Dim hImage As Integer
		Dim XExt As Integer
		Dim YExt As Integer
	End Structure
	Private Structure CLSID
		Dim Data1 As Integer
		Dim Data2 As Short
		Dim Data3 As Short
		<VBFixedArray(7)> Dim Data4() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim Data4(7)
		End Sub
	End Structure
	Private Structure FILETIME
		Dim dwLowDateTime As Integer
		Dim dwHighDateTime As Integer
	End Structure
	Private Structure SYSTEMTIME
		Dim wYear As Short
		Dim wMonth As Short
		Dim wDayOfWeek As Short
		Dim wDay As Short
		Dim wHour As Short
		Dim wMinute As Short
		Dim wSecond As Short
		Dim wMilliseconds As Short
	End Structure
	Private Const MAX_PATH As Integer = 260
	Private Structure WIN32_FIND_DATA
		Dim dwFileAttributes As Integer
		Dim FTCreationTime As FILETIME
		Dim FTLastAccessTime As FILETIME
		Dim FTLastWriteTime As FILETIME
		Dim nFileSizeHigh As Integer
		Dim nFileSizeLow As Integer
		Dim dwReserved0 As Integer
		Dim dwReserved1 As Integer
		<VBFixedArray(((MAX_PATH * 2) - 1))> Dim lpszFileName() As Byte
		<VBFixedArray(((14 * 2) - 1))> Dim lpszAlternateFileName() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim lpszFileName(((MAX_PATH * 2) - 1))
			ReDim lpszAlternateFileName(((14 * 2) - 1))
		End Sub
	End Structure
	Private Structure VS_FIXEDFILEINFO
		Dim dwSignature As Integer
		Dim dwStrucVersionLo As Short
		Dim dwStrucVersionHi As Short
		Dim dwFileVersionMSLo As Short
		Dim dwFileVersionMSHi As Short
		Dim dwFileVersionLSLo As Short
		Dim dwFileVersionLSHi As Short
		Dim dwProductVersionMSLo As Short
		Dim dwProductVersionMSHi As Short
		Dim dwProductVersionLSLo As Short
		Dim dwProductVersionLSHi As Short
		Dim dwFileFlagsMask As Integer
		Dim dwFileFlags As Integer
		Dim dwFileOS As Integer
		Dim dwFileType As Integer
		Dim dwFileSubtype As Integer
		Dim dwFileDateMS As Integer
		Dim dwFileDateLS As Integer
	End Structure
	Private Structure MONITORINFO
		Dim cbSize As Integer
		Dim RCMonitor As RECT
		Dim RCWork As RECT
		Dim dwFlags As Integer
	End Structure
	Private Structure FLASHWINFO
		Dim cbSize As Integer
		Dim hWnd As Integer
		Dim dwFlags As Integer
		Dim uCount As Integer
		Dim dwTimeout As Integer
	End Structure
	Private Const LF_FACESIZE As Integer = 32
	Private Const DEFAULT_QUALITY As Integer = 0
	Private Structure LOGFONT
		Dim LFHeight As Integer
		Dim LFWidth As Integer
		Dim LFEscapement As Integer
		Dim LFOrientation As Integer
		Dim LFWeight As Integer
		Dim LFItalic As Byte
		Dim LFUnderline As Byte
		Dim LFStrikeOut As Byte
		Dim LFCharset As Byte
		Dim LFOutPrecision As Byte
		Dim LFClipPrecision As Byte
		Dim LFQuality As Byte
		Dim LFPitchAndFamily As Byte
		<VBFixedArray(((LF_FACESIZE * 2) - 1))> Dim LFFaceName() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim LFFaceName(((LF_FACESIZE * 2) - 1))
		End Sub
	End Structure
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef Destination As Any, ByRef Source As Any, ByVal Length As Integer)
	Private Declare Function ArrPtr Lib "msvbvm60.dll"  Alias "VarPtr"(ByRef Var() As Any) As Integer
	Private Declare Function lstrlen Lib "kernel32"  Alias "lstrlenW"(ByVal lpString As Integer) As Integer
	Private Declare Function lstrcpy Lib "kernel32"  Alias "lstrcpyW"(ByVal lpString1 As Integer, ByVal lpString2 As Integer) As Integer
	'UPGRADE_WARNING: Structure MSGBOXPARAMS may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function MessageBoxIndirect Lib "user32"  Alias "MessageBoxIndirectW"(ByRef lpMsgBoxParams As MSGBOXPARAMS) As Integer
	Private Declare Function GetActiveWindow Lib "user32" () As Integer
	Private Declare Function GetForegroundWindow Lib "user32" () As Integer
	Private Declare Function GetFileAttributes Lib "kernel32"  Alias "GetFileAttributesW"(ByVal lpFileName As Integer) As Integer
	Private Declare Function SetFileAttributes Lib "kernel32"  Alias "SetFileAttributesW"(ByVal lpFileName As Integer, ByVal dwFileAttributes As Integer) As Integer
	Private Declare Function CreateFile Lib "kernel32"  Alias "CreateFileW"(ByVal lpFileName As Integer, ByVal dwDesiredAccess As Integer, ByVal dwShareMode As Integer, ByVal lpSecurityAttributes As Integer, ByVal dwCreationDisposition As Integer, ByVal dwFlagsAndAttributes As Integer, ByVal hTemplateFile As Integer) As Integer
	Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Integer, ByRef lpFileSizeHigh As Integer) As Integer
	Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Integer, ByVal lpCreationTime As Integer, ByVal lpLastAccessTime As Integer, ByVal lpLastWriteTime As Integer) As Integer
	Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (ByVal lpFileTime As Integer, ByVal lpLocalFileTime As Integer) As Integer
	Private Declare Function FileTimeToSystemTime Lib "kernel32" (ByVal lpFileTime As Integer, ByVal lpSystemTime As Integer) As Integer
	'UPGRADE_WARNING: Structure WIN32_FIND_DATA may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function FindFirstFile Lib "kernel32"  Alias "FindFirstFileW"(ByVal lpFileName As Integer, ByRef lpFindFileData As WIN32_FIND_DATA) As Integer
	'UPGRADE_WARNING: Structure WIN32_FIND_DATA may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function FindNextFile Lib "kernel32"  Alias "FindNextFileW"(ByVal hFindFile As Integer, ByRef lpFindFileData As WIN32_FIND_DATA) As Integer
	Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Integer, ByRef lpRect As RECT) As Integer
	Private Declare Function MonitorFromWindow Lib "user32" (ByVal hWnd As Integer, ByVal dwFlags As Integer) As Integer
	'UPGRADE_WARNING: Structure MONITORINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetMonitorInfo Lib "user32"  Alias "GetMonitorInfoW"(ByVal hMonitor As Integer, ByRef lpMI As MONITORINFO) As Integer
	Private Declare Function GetVolumePathName Lib "kernel32"  Alias "GetVolumePathNameW"(ByVal lpFileName As Integer, ByVal lpVolumePathName As Integer, ByVal cch As Integer) As Integer
	Private Declare Function GetVolumeInformation Lib "kernel32"  Alias "GetVolumeInformationW"(ByVal lpRootPathName As Integer, ByVal lpVolumeNameBuffer As Integer, ByVal nVolumeNameSize As Integer, ByRef lpVolumeSerialNumber As Integer, ByRef lpMaximumComponentLength As Integer, ByRef lpFileSystemFlags As Integer, ByVal lpFileSystemNameBuffer As Integer, ByVal nFileSystemNameSize As Integer) As Integer
	Private Declare Function CreateDirectory Lib "kernel32"  Alias "CreateDirectoryW"(ByVal lpPathName As Integer, ByVal lpSecurityAttributes As Integer) As Integer
	Private Declare Function RemoveDirectory Lib "kernel32"  Alias "RemoveDirectoryW"(ByVal lpPathName As Integer) As Integer
	Private Declare Function GetFileVersionInfo Lib "Version"  Alias "GetFileVersionInfoW"(ByVal lpFileName As Integer, ByVal dwHandle As Integer, ByVal dwLen As Integer, ByVal lpData As Integer) As Integer
	Private Declare Function GetFileVersionInfoSize Lib "Version"  Alias "GetFileVersionInfoSizeW"(ByVal lpFileName As Integer, ByVal lpdwHandle As Integer) As Integer
	Private Declare Function VerQueryValue Lib "Version"  Alias "VerQueryValueW"(ByVal lpBlock As Integer, ByVal lpSubBlock As Integer, ByRef lplpBuffer As Integer, ByRef puLen As Integer) As Integer
	Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
	Private Declare Function GetCommandLine Lib "kernel32"  Alias "GetCommandLineW"() As Integer
	Private Declare Function PathGetArgs Lib "shlwapi"  Alias "PathGetArgsW"(ByVal lpszPath As Integer) As Integer
	Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pbString As Integer, ByVal pszStrPtr As Integer) As Integer
	Private Declare Function VarDecFromI8 Lib "oleaut32" (ByVal LoDWord As Integer, ByVal HiDWord As Integer, ByRef pDecOut As Object) As Integer
	Private Declare Function GetModuleFileName Lib "kernel32"  Alias "GetModuleFileNameW"(ByVal hModule As Integer, ByVal lpFileName As Integer, ByVal nSize As Integer) As Integer
	Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Integer) As Integer
	Private Declare Function EmptyClipboard Lib "user32" () As Integer
	Private Declare Function CloseClipboard Lib "user32" () As Integer
	Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Integer
	Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Integer
	Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Integer, ByVal hMem As Integer) As Integer
	Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Integer) As Short
	Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short
	Private Declare Function GetWindowText Lib "user32"  Alias "GetWindowTextW"(ByVal hWnd As Integer, ByVal lpString As Integer, ByVal cch As Integer) As Integer
	Private Declare Function GetWindowTextLength Lib "user32"  Alias "GetWindowTextLengthW"(ByVal hWnd As Integer) As Integer
	Private Declare Function GetClassName Lib "user32"  Alias "GetClassNameW"(ByVal hWnd As Integer, ByVal lpClassName As Integer, ByVal nMaxCount As Integer) As Integer
	Private Declare Function GetSystemWindowsDirectory Lib "kernel32"  Alias "GetSystemWindowsDirectoryW"(ByVal lpBuffer As Integer, ByVal nSize As Integer) As Integer
	Private Declare Function GetSystemDirectory Lib "kernel32"  Alias "GetSystemDirectoryW"(ByVal lpBuffer As Integer, ByVal nSize As Integer) As Integer
	Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Integer) As Integer
	Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Integer) As Integer
	'UPGRADE_WARNING: Structure POINTAPI may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Integer
	Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Integer, ByVal Y As Integer) As Integer
	Private Declare Function GetCapture Lib "user32" () As Integer
	Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Integer, ByVal lpdwProcessId As Integer) As Integer
	'UPGRADE_WARNING: Structure FLASHWINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function FlashWindowEx Lib "user32" (ByRef pFWI As FLASHWINFO) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function SendMessage Lib "user32"  Alias "SendMessageW"(ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Any) As Integer
	Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Integer, ByVal lprcUpdate As Integer, ByVal hrgnUpdate As Integer, ByVal fuRedraw As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function GetObjectAPI Lib "gdi32"  Alias "GetObjectW"(ByVal hObject As Integer, ByVal nCount As Integer, ByRef lpObject As Any) As Integer
	Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
	Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Integer) As Integer
	Private Declare Function GetDC Lib "user32" (ByVal hWnd As Integer) As Integer
	Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Integer, ByVal nIndex As Integer) As Integer
	Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Integer, ByVal hDC As Integer) As Integer
	Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Integer) As Integer
	Private Declare Function GdiAlphaBlend Lib "gdi32" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal nWidthSrc As Integer, ByVal nHeightSrc As Integer, ByVal BlendFunc As Integer) As Integer
	Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Integer, ByVal XLeft As Integer, ByVal YTop As Integer, ByVal hIcon As Integer, ByVal CXWidth As Integer, ByVal CYWidth As Integer, ByVal istepIfAniCur As Integer, ByVal hbrFlickerFreeDraw As Integer, ByVal diFlags As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function FillRect Lib "user32" (ByVal hDC As Integer, ByRef lpRect As RECT, ByVal hBrush As Integer) As Integer
	Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Integer) As Integer
	Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Integer) As Integer
	Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
	Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Integer, ByVal nNumerator As Integer, ByVal nDenominator As Integer) As Integer
	'UPGRADE_WARNING: Structure LOGFONT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function CreateFontIndirect Lib "gdi32"  Alias "CreateFontIndirectW"(ByRef lpLogFont As LOGFONT) As Integer
	Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Integer, ByVal dwBytes As Integer) As Integer
	Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Integer) As Integer
	Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Integer) As Integer
	Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Integer) As Integer
	Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal Color As Integer, ByVal hPal As Integer, ByRef RGBResult As Integer) As Integer
	'UPGRADE_WARNING: Structure IPicture may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_WARNING: Structure IUnknown may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function OleLoadPicture Lib "oleaut32" (ByVal pStream As stdole.IUnknown, ByVal lSize As Integer, ByVal fRunmode As Integer, ByRef riid As Any, ByRef pIPicture As System.Drawing.Image) As Integer
	'UPGRADE_WARNING: Structure IPicture may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure CLSID may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure OLE_COLOR may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function OleLoadPicturePath Lib "oleaut32" (ByVal lpszPath As Integer, ByVal pUnkCaller As Integer, ByVal dwReserved As Integer, ByVal ClrReserved As System.Drawing.Color, ByRef riid As CLSID, ByRef pIPicture As System.Drawing.Image) As Integer
	'UPGRADE_WARNING: Structure IPicture may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_WARNING: Structure PICTDESC may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function OleCreatePictureIndirect Lib "olepro32" (ByRef pPictDesc As PICTDESC, ByRef riid As Any, ByVal fPictureOwnsHandle As Integer, ByRef pIPicture As System.Drawing.Image) As Integer
	'UPGRADE_WARNING: Structure IUnknown may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Integer, ByVal fDeleteOnRelease As Integer, ByRef pStream As stdole.IUnknown) As Integer
	Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Integer, ByVal dwFlags As Integer, ByVal lpWideCharStr As Integer, ByVal cchWideChar As Integer, ByVal lpMultiByteStr As Integer, ByVal cbMultiByte As Integer, ByVal lpDefaultChar As Integer, ByVal lpUsedDefaultChar As Integer) As Integer
	Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Integer, ByVal dwFlags As Integer, ByVal lpMultiByteStr As Integer, ByVal cbMultiByte As Integer, ByVal lpWideCharStr As Integer, ByVal cchWideChar As Integer) As Integer
	
	' (VB-Overwrite)
	'UPGRADE_NOTE: MsgBox was upgraded to MsgBox_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function MsgBox_Renamed(ByVal Prompt As String, Optional ByVal Buttons As MsgBoxStyle = MsgBoxStyle.OKOnly, Optional ByVal Title As String = "") As MsgBoxResult
		Dim MSGBOXP As MSGBOXPARAMS
		With MSGBOXP
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			.cbSize = LenB(MSGBOXP)
			If (Buttons And MsgBoxStyle.SystemModal) = 0 Then
				If Not System.Windows.Forms.Form.ActiveForm Is Nothing Then
					.hWndOwner = System.Windows.Forms.Form.ActiveForm.Handle.ToInt32
				Else
					.hWndOwner = GetActiveWindow()
				End If
			Else
				.hWndOwner = GetForegroundWindow()
			End If
			.hInstance = VB6.GetHInstance.ToInt32
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			.lpszText = StrPtr(Prompt)
			If Title = vbNullString Then Title = My.Application.Info.Title
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			.lpszCaption = StrPtr(Title)
			.dwStyle = Buttons
		End With
		MsgBox_Renamed = MessageBoxIndirect(MSGBOXP)
	End Function
	
	' (VB-Overwrite)
	Public Sub SendKeys(ByRef Text As String, Optional ByRef Wait As Boolean = False)
		'UPGRADE_WARNING: Couldn't resolve default property of object CreateObject().SendKeys. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CreateObject("WScript.Shell").SendKeys(Text, Wait)
	End Sub
	
	' (VB-Overwrite)
	'UPGRADE_NOTE: GetAttr was upgraded to GetAttr_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function GetAttr_Renamed(ByVal PathName As String) As FileAttribute
		Const INVALID_FILE_ATTRIBUTES As Integer = (-1)
		Const FILE_ATTRIBUTE_NORMAL As Integer = &H80
		If Left(PathName, 2) = "\\" Then PathName = "UNC\" & Mid(PathName, 3)
		Dim dwAttributes As Integer
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		dwAttributes = GetFileAttributes(StrPtr("\\?\" & PathName))
		If dwAttributes = INVALID_FILE_ATTRIBUTES Then
			Err.Raise(53)
		ElseIf dwAttributes = FILE_ATTRIBUTE_NORMAL Then 
			'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
			GetAttr_Renamed = vbNormal
		Else
			GetAttr_Renamed = dwAttributes
		End If
	End Function
	
	' (VB-Overwrite)
	'UPGRADE_NOTE: SetAttr was upgraded to SetAttr_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub SetAttr_Renamed(ByVal PathName As String, ByVal Attributes As FileAttribute)
		Const FILE_ATTRIBUTE_NORMAL As Integer = &H80
		Dim dwAttributes As Integer
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		If Attributes = vbNormal Then
			dwAttributes = FILE_ATTRIBUTE_NORMAL
		Else
			'UPGRADE_ISSUE: Constant vbAlias was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
			If (Attributes And (FileAttribute.Volume Or FileAttribute.Directory Or vbAlias)) <> 0 Then Err.Raise(5)
			dwAttributes = Attributes
		End If
		If Left(PathName, 2) = "\\" Then PathName = "UNC\" & Mid(PathName, 3)
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If SetFileAttributes(StrPtr("\\?\" & PathName), dwAttributes) = 0 Then Err.Raise(53)
	End Sub
	
	' (VB-Overwrite)
	'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
	'UPGRADE_NOTE: Dir was upgraded to Dir_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function Dir_Renamed(Optional ByVal PathMask As String = "", Optional ByVal Attributes As FileAttribute = vbNormal) As String
		Const INVALID_HANDLE_VALUE As Integer = (-1)
		Const FILE_ATTRIBUTE_NORMAL As Integer = &H80
		Static hFindFile As Integer
		Static AttributesCache As FileAttribute
		Dim VolumePathBuffer, VolumeNameBuffer As String
		'UPGRADE_WARNING: Arrays in structure FD may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim FD As WIN32_FIND_DATA
		Dim dwMask As Integer
		If Attributes = FileAttribute.Volume Then ' Exact match
			' If any other attribute is specified, vbVolume is ignored.
			If hFindFile <> 0 Then
				FindClose(hFindFile)
				hFindFile = 0
			End If
			If Len(PathMask) = 0 Then
				VolumeNameBuffer = New String(vbNullChar, MAX_PATH)
				'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				If GetVolumeInformation(0, StrPtr(VolumeNameBuffer), Len(VolumeNameBuffer), 0, 0, 0, 0, 0) <> 0 Then Dir_Renamed = Left(VolumeNameBuffer, InStr(VolumeNameBuffer, vbNullChar) - 1)
			Else
				VolumePathBuffer = New String(vbNullChar, MAX_PATH)
				If Left(PathMask, 2) = "\\" Then PathMask = "UNC\" & Mid(PathMask, 3)
				'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				If GetVolumePathName(StrPtr("\\?\" & PathMask), StrPtr(VolumePathBuffer), Len(VolumePathBuffer)) <> 0 Then
					VolumePathBuffer = Left(VolumePathBuffer, InStr(VolumePathBuffer, vbNullChar) - 1)
					VolumeNameBuffer = New String(vbNullChar, MAX_PATH)
					'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					If GetVolumeInformation(StrPtr(VolumePathBuffer), StrPtr(VolumeNameBuffer), Len(VolumeNameBuffer), 0, 0, 0, 0, 0) <> 0 Then Dir_Renamed = Left(VolumeNameBuffer, InStr(VolumeNameBuffer, vbNullChar) - 1)
				End If
			End If
		Else
			If Len(PathMask) = 0 Then
				If hFindFile <> 0 Then
					If FindNextFile(hFindFile, FD) = 0 Then
						FindClose(hFindFile)
						hFindFile = 0
						Exit Function
					End If
				Else
					Err.Raise(5)
					Exit Function
				End If
			Else
				If hFindFile <> 0 Then
					FindClose(hFindFile)
					hFindFile = 0
				End If
				Select Case Right(PathMask, 1)
					Case "\", ":", "/"
						PathMask = PathMask & "*.*"
				End Select
				AttributesCache = Attributes
				If Left(PathMask, 2) = "\\" Then PathMask = "UNC\" & Mid(PathMask, 3)
				'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				hFindFile = FindFirstFile(StrPtr("\\?\" & PathMask), FD)
				If hFindFile = INVALID_HANDLE_VALUE Then
					hFindFile = 0
					If Err.LastDllError > 12 Then Err.Raise(52)
					Exit Function
				End If
			End If
			Do 
				If FD.dwFileAttributes = FILE_ATTRIBUTE_NORMAL Then
					dwMask = 0 ' Found
				Else
					dwMask = FD.dwFileAttributes And (Not AttributesCache) And &H16
				End If
				If dwMask = 0 Then
					Dir_Renamed = Left(System.Text.UnicodeEncoding.Unicode.GetString(FD.lpszFileName), InStr(System.Text.UnicodeEncoding.Unicode.GetString(FD.lpszFileName), vbNullChar) - 1)
					If FD.dwFileAttributes And FileAttribute.Directory Then
						If Dir_Renamed <> "." And Dir_Renamed <> ".." Then Exit Do ' Exclude self and relative path aliases
					Else
						Exit Do
					End If
				End If
				If FindNextFile(hFindFile, FD) = 0 Then
					FindClose(hFindFile)
					hFindFile = 0
					Exit Do
				End If
			Loop 
		End If
	End Function
	
	' (VB-Overwrite)
	'UPGRADE_NOTE: MkDir was upgraded to MkDir_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub MkDir_Renamed(ByVal PathName As String)
		If Left(PathName, 2) = "\\" Then PathName = "UNC\" & Mid(PathName, 3)
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		Const ERROR_PATH_NOT_FOUND As Integer = 3
		If CreateDirectory(StrPtr("\\?\" & PathName), 0) = 0 Then
			If Err.LastDllError = ERROR_PATH_NOT_FOUND Then
				Err.Raise(76)
			Else
				Err.Raise(75)
			End If
		End If
	End Sub
	
	' (VB-Overwrite)
	'UPGRADE_NOTE: RmDir was upgraded to RmDir_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub RmDir_Renamed(ByVal PathName As String)
		If Left(PathName, 2) = "\\" Then PathName = "UNC\" & Mid(PathName, 3)
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		Const ERROR_FILE_NOT_FOUND As Integer = 2
		If RemoveDirectory(StrPtr("\\?\" & PathName)) = 0 Then
			If Err.LastDllError = ERROR_FILE_NOT_FOUND Then
				Err.Raise(76)
			Else
				Err.Raise(75)
			End If
		End If
	End Sub
	
	' (VB-Overwrite)
	'UPGRADE_NOTE: FileLen was upgraded to FileLen_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function FileLen_Renamed(ByVal PathName As String) As Object
		Const INVALID_HANDLE_VALUE As Integer = (-1)
		Const INVALID_FILE_SIZE As Integer = (-1)
		Const GENERIC_READ As Integer = &H80000000
		Const FILE_SHARE_READ As Integer = &H1
		Const OPEN_EXISTING As Integer = 3
		Const FILE_FLAG_SEQUENTIAL_SCAN As Integer = &H8000000
		Dim hFile As Integer
		If Left(PathName, 2) = "\\" Then PathName = "UNC\" & Mid(PathName, 3)
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		hFile = CreateFile(StrPtr("\\?\" & PathName), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
		Dim LoDWord, HiDWord As Integer
		If hFile <> INVALID_HANDLE_VALUE Then
			LoDWord = GetFileSize(hFile, HiDWord)
			CloseHandle(hFile)
			If LoDWord <> INVALID_FILE_SIZE Then
				FileLen_Renamed = CDec(0)
				VarDecFromI8(LoDWord, HiDWord, FileLen_Renamed)
			Else
				'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
				'UPGRADE_WARNING: Couldn't resolve default property of object FileLen_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				FileLen_Renamed = System.DBNull.Value
			End If
		Else
			Err.Raise(Number:=53, Description:="File not found: '" & PathName & "'")
		End If
	End Function
	
	' (VB-Overwrite)
	'UPGRADE_NOTE: FileDateTime was upgraded to FileDateTime_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function FileDateTime_Renamed(ByVal PathName As String) As Date
		Const INVALID_HANDLE_VALUE As Integer = (-1)
		Const GENERIC_READ As Integer = &H80000000
		Const FILE_SHARE_READ As Integer = &H1
		Const OPEN_EXISTING As Integer = 3
		Const FILE_FLAG_SEQUENTIAL_SCAN As Integer = &H8000000
		Dim hFile As Integer
		If Left(PathName, 2) = "\\" Then PathName = "UNC\" & Mid(PathName, 3)
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		hFile = CreateFile(StrPtr("\\?\" & PathName), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
		Dim FT(1) As FILETIME
		Dim ST As SYSTEMTIME
		If hFile <> INVALID_HANDLE_VALUE Then
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			GetFileTime(hFile, 0, 0, VarPtr(FT(0)))
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			FileTimeToLocalFileTime(VarPtr(FT(0)), VarPtr(FT(1)))
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			FileTimeToSystemTime(VarPtr(FT(1)), VarPtr(ST))
			FileDateTime_Renamed = System.Date.FromOADate(DateSerial(ST.wYear, ST.wMonth, ST.wDay).ToOADate + TimeSerial(ST.wHour, ST.wMinute, ST.wSecond).ToOADate)
			CloseHandle(hFile)
		Else
			Err.Raise(Number:=53, Description:="File not found: '" & PathName & "'")
		End If
	End Function
	
	' (VB-Overwrite)
	'UPGRADE_NOTE: Command was upgraded to Command_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function Command_Renamed() As String
		If InIDE() = False Then
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			SysReAllocString(VarPtr(Command_Renamed), PathGetArgs(GetCommandLine()))
			Command_Renamed = LTrim(Command_Renamed)
		Else
			Command_Renamed = VB.Command()
		End If
	End Function
	
	Public Function FileExists(ByVal PathName As String) As Boolean
		On Error Resume Next
		Dim Attributes As FileAttribute
		Dim ErrVal As Integer
		Attributes = GetAttr_Renamed(PathName)
		ErrVal = Err.Number
		On Error GoTo 0
		If (Attributes And (FileAttribute.Directory Or FileAttribute.Volume)) = 0 And ErrVal = 0 Then FileExists = True
	End Function
	
	Public Function AppPath() As String
		Const MAX_PATH_W As Integer = 32767
		Dim Buffer As String
		Dim RetVal As Integer
		If InIDE() = False Then
			Buffer = New String(vbNullChar, MAX_PATH)
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			RetVal = GetModuleFileName(0, StrPtr(Buffer), MAX_PATH)
			If RetVal = MAX_PATH Then ' Path > MAX_PATH
				Buffer = New String(vbNullChar, MAX_PATH_W)
				'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				RetVal = GetModuleFileName(0, StrPtr(Buffer), MAX_PATH_W)
			End If
			If RetVal > 0 Then
				Buffer = Left(Buffer, RetVal)
				AppPath = Left(Buffer, InStrRev(Buffer, "\"))
			Else
				AppPath = My.Application.Info.DirectoryPath & IIf(Right(My.Application.Info.DirectoryPath, 1) = "\", "", "\")
			End If
		Else
			AppPath = My.Application.Info.DirectoryPath & IIf(Right(My.Application.Info.DirectoryPath, 1) = "\", "", "\")
		End If
	End Function
	
	Public Function AppEXEName() As String
		Const MAX_PATH_W As Integer = 32767
		Dim Buffer As String
		Dim RetVal As Integer
		If InIDE() = False Then
			Buffer = New String(vbNullChar, MAX_PATH)
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			RetVal = GetModuleFileName(0, StrPtr(Buffer), MAX_PATH)
			If RetVal = MAX_PATH Then ' Path > MAX_PATH
				Buffer = New String(vbNullChar, MAX_PATH_W)
				'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				RetVal = GetModuleFileName(0, StrPtr(Buffer), MAX_PATH_W)
			End If
			If RetVal > 0 Then
				Buffer = Left(Buffer, RetVal)
				Buffer = Right(Buffer, Len(Buffer) - InStrRev(Buffer, "\"))
				AppEXEName = Left(Buffer, InStrRev(Buffer, ".") - 1)
			Else
				'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				AppEXEName = My.Application.Info.AssemblyName
			End If
		Else
			'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			AppEXEName = My.Application.Info.AssemblyName
		End If
	End Function
	
	Public Function AppMajor() As Short
		If InIDE() = False Then
			With GetAppVersionInfo()
				AppMajor = .dwFileVersionMSHi
			End With
		Else
			AppMajor = My.Application.Info.Version.Major
		End If
	End Function
	
	Public Function AppMinor() As Short
		If InIDE() = False Then
			With GetAppVersionInfo()
				AppMinor = .dwFileVersionMSLo
			End With
		Else
			AppMinor = My.Application.Info.Version.Minor
		End If
	End Function
	
	Public Function AppRevision() As Short
		If InIDE() = False Then
			With GetAppVersionInfo()
				AppRevision = .dwFileVersionLSLo
			End With
		Else
			AppRevision = My.Application.Info.Version.Revision
		End If
	End Function
	
	Private Function GetAppVersionInfo() As VS_FIXEDFILEINFO
		Static Done As Boolean
		Static Value As VS_FIXEDFILEINFO
		Const MAX_PATH_W As Integer = 32767
		Dim DataBuffer() As Byte
		Dim hData As Integer
		Dim ImagePath As String
		Dim Length As Integer
		Dim Buffer As String
		Dim RetVal As Integer
		If Done = False Then
			Buffer = New String(vbNullChar, MAX_PATH)
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			RetVal = GetModuleFileName(0, StrPtr(Buffer), MAX_PATH)
			If RetVal = MAX_PATH Then ' Path > MAX_PATH
				Buffer = New String(vbNullChar, MAX_PATH_W)
				'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				RetVal = GetModuleFileName(0, StrPtr(Buffer), MAX_PATH_W)
			End If
			If RetVal > 0 Then
				ImagePath = Left(Buffer, RetVal)
				'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				Length = GetFileVersionInfoSize(StrPtr(ImagePath), 0)
				If Length > 0 Then
					ReDim DataBuffer((Length - 1))
					'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					If GetFileVersionInfo(StrPtr(ImagePath), 0, Length, VarPtr(DataBuffer(0))) <> 0 Then
						'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
						'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
						If VerQueryValue(VarPtr(DataBuffer(0)), StrPtr("\"), hData, Length) <> 0 Then
							'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If hData <> 0 Then CopyMemory(Value, hData, LenB(Value))
						End If
					End If
				End If
			End If
			Done = True
		End If
		'UPGRADE_ISSUE: LSet cannot assign one type to another. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="899FA812-8F71-4014-BAEE-E5AF348BA5AA"'
		GetAppVersionInfo = LSet(Value)
	End Function
	
	Public Function GetClipboardText() As String
		Const CF_UNICODETEXT As Integer = 13
		Dim lpMem, lpText, Length As Integer
		If OpenClipboard(0) <> 0 Then
			If IsClipboardFormatAvailable(CF_UNICODETEXT) <> 0 Then
				lpText = GetClipboardData(CF_UNICODETEXT)
				If lpText <> 0 Then
					lpMem = GlobalLock(lpText)
					If lpMem <> 0 Then
						Length = lstrlen(lpMem)
						If Length > 0 Then
							GetClipboardText = New String(vbNullChar, Length)
							'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
							lstrcpy(StrPtr(GetClipboardText), lpMem)
						End If
						GlobalUnlock(lpMem)
					End If
				End If
			End If
			CloseClipboard()
		End If
	End Function
	
	Public Sub SetClipboardText(ByRef Text As String)
		Const CF_UNICODETEXT As Integer = 13
		Const GMEM_MOVEABLE As Integer = &H2
		Dim Buffer As String
		Dim Length As Integer
		Dim hMem, lpMem As Integer
		If OpenClipboard(0) <> 0 Then
			EmptyClipboard()
			Buffer = Text & vbNullChar
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			Length = LenB(Buffer)
			hMem = GlobalAlloc(GMEM_MOVEABLE, Length)
			If hMem <> 0 Then
				lpMem = GlobalLock(hMem)
				If lpMem <> 0 Then
					'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					CopyMemory(lpMem, StrPtr(Buffer), Length)
					GlobalUnlock(hMem)
					SetClipboardData(CF_UNICODETEXT, hMem)
				End If
			End If
			CloseClipboard()
		End If
	End Sub
	
	Public Function AccelCharCode(ByVal Caption As String) As Short
		If Caption = vbNullString Then Exit Function
		Dim Pos, Length As Integer
		Length = Len(Caption)
		Pos = Length
		Do 
			If Mid(Caption, Pos, 1) = "&" And Pos < Length Then
				AccelCharCode = Asc(UCase(Mid(Caption, Pos + 1, 1)))
				If Pos > 1 Then
					If Mid(Caption, Pos - 1, 1) = "&" Then AccelCharCode = 0
				Else
					If AccelCharCode = System.Windows.Forms.Keys.Up Then AccelCharCode = 0
				End If
				If AccelCharCode <> 0 Then Exit Do
			End If
			Pos = Pos - 1
		Loop Until Pos = 0
	End Function
	
	Public Function ProperControlName(ByVal Control As System.Windows.Forms.Control) As String
		Dim Index As Integer
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object Control.Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Index = Control.Index
		If Err.Number <> 0 Or Index < 0 Then ProperControlName = Control.Name Else ProperControlName = Control.Name & "(" & Index & ")"
		On Error GoTo 0
	End Function
	
	Public Function GetTopUserControl(ByVal UserControl As Object) As System.Windows.Forms.UserControl
		If UserControl Is Nothing Then Exit Function
		Dim TopUserControl, TempUserControl As System.Windows.Forms.UserControl
		'UPGRADE_ISSUE: ObjPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object TempUserControl.Controls.Item. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(TempUserControl.Controls.Item, ObjPtr(UserControl), 4)
		TopUserControl = TempUserControl
		'UPGRADE_WARNING: Couldn't resolve default property of object TempUserControl.Controls.Item. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(TempUserControl.Controls.Item, 0, 4)
		'UPGRADE_ISSUE: VBRUN.ParentControlsType object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim TempParentControlsType As VBRUN.ParentControlsType
		'UPGRADE_ISSUE: VBRUN.ParentControlsType object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim OldParentControlsType As VBRUN.ParentControlsType
		With TopUserControl
			'UPGRADE_ISSUE: ParentControls property ParentControls.Count was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
			If .Parent.Controls.Count > 0 Then
				'UPGRADE_ISSUE: ParentControls property ParentControls.ParentControlsType was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
				OldParentControlsType = .Parent.Controls.ParentControlsType
				'UPGRADE_ISSUE: ParentControls property ParentControls.ParentControlsType was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
				'UPGRADE_ISSUE: Constant vbExtender was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
				.Parent.Controls.ParentControlsType = vbExtender
				If TypeOf .Parent.Controls(0) Is System.Windows.Forms.AxHost Then
					'UPGRADE_ISSUE: ParentControls property ParentControls.ParentControlsType was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
					'UPGRADE_ISSUE: Constant vbNoExtender was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
					.Parent.Controls.ParentControlsType = vbNoExtender
					'UPGRADE_ISSUE: ObjPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object TempUserControl.Controls.Item. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CopyMemory(TempUserControl.Controls.Item, ObjPtr(.Parent.Controls(0)), 4)
					TopUserControl = TempUserControl
					'UPGRADE_WARNING: Couldn't resolve default property of object TempUserControl.Controls.Item. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CopyMemory(TempUserControl.Controls.Item, 0, 4)
					Do 
						With TopUserControl
							'UPGRADE_ISSUE: ParentControls property ParentControls.Count was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
							If .Parent.Controls.Count = 0 Then Exit Do
							'UPGRADE_ISSUE: ParentControls property ParentControls.ParentControlsType was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
							TempParentControlsType = .Parent.Controls.ParentControlsType
							'UPGRADE_ISSUE: ParentControls property ParentControls.ParentControlsType was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
							'UPGRADE_ISSUE: Constant vbExtender was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
							.Parent.Controls.ParentControlsType = vbExtender
							If TypeOf .Parent.Controls(0) Is System.Windows.Forms.AxHost Then
								'UPGRADE_ISSUE: ParentControls property ParentControls.ParentControlsType was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
								'UPGRADE_ISSUE: Constant vbNoExtender was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
								.Parent.Controls.ParentControlsType = vbNoExtender
								'UPGRADE_ISSUE: ObjPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object TempUserControl.Controls.Item. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								CopyMemory(TempUserControl.Controls.Item, ObjPtr(.Parent.Controls(0)), 4)
								TopUserControl = TempUserControl
								'UPGRADE_WARNING: Couldn't resolve default property of object TempUserControl.Controls.Item. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								CopyMemory(TempUserControl.Controls.Item, 0, 4)
								'UPGRADE_ISSUE: ParentControls property ParentControls.ParentControlsType was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
								.Parent.Controls.ParentControlsType = TempParentControlsType
							Else
								'UPGRADE_ISSUE: ParentControls property ParentControls.ParentControlsType was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
								.Parent.Controls.ParentControlsType = TempParentControlsType
								Exit Do
							End If
						End With
					Loop 
				End If
				'UPGRADE_ISSUE: ParentControls property ParentControls.ParentControlsType was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
				.Parent.Controls.ParentControlsType = OldParentControlsType
			End If
		End With
		GetTopUserControl = TopUserControl
	End Function
	
	Public Function MousePointerID(ByVal MousePointer As Short) As Integer
		Const IDC_ARROW As Integer = 32512
		Const IDC_CROSS As Integer = 32515
		Const IDC_IBEAM As Integer = 32513
		Const IDC_HAND As Integer = 32649
		Const IDC_SIZEALL As Integer = 32646
		Const IDC_SIZENESW As Integer = 32643
		Const IDC_SIZENS As Integer = 32645
		Const IDC_SIZENWSE As Integer = 32642
		Const IDC_SIZEWE As Integer = 32644
		Const IDC_UPARROW As Integer = 32516
		Const IDC_WAIT As Integer = 32514
		Const IDC_NO As Integer = 32648
		Const IDC_APPSTARTING As Integer = 32650
		Const IDC_HELP As Integer = 32651
		Const IDC_WAITCD As Integer = 32663 ' Undocumented
		Select Case MousePointer
			Case System.Windows.Forms.Cursors.Arrow
				MousePointerID = IDC_ARROW
			Case System.Windows.Forms.Cursors.Cross
				MousePointerID = IDC_CROSS
			Case System.Windows.Forms.Cursors.IBeam
				MousePointerID = IDC_IBEAM
			Case System.Windows.Forms.Cursors.Default ' Obselete, replaced Icon with Hand
				MousePointerID = IDC_HAND
			Case System.Windows.Forms.Cursors.SizeAll, System.Windows.Forms.Cursors.SizeAll
				MousePointerID = IDC_SIZEALL
			Case System.Windows.Forms.Cursors.SizeNESW
				MousePointerID = IDC_SIZENESW
			Case System.Windows.Forms.Cursors.SizeNS
				MousePointerID = IDC_SIZENS
			Case System.Windows.Forms.Cursors.SizeNWSE
				MousePointerID = IDC_SIZENWSE
			Case System.Windows.Forms.Cursors.SizeWE
				MousePointerID = IDC_SIZEWE
			Case System.Windows.Forms.Cursors.UpArrow
				MousePointerID = IDC_UPARROW
			Case System.Windows.Forms.Cursors.WaitCursor
				MousePointerID = IDC_WAIT
			Case System.Windows.Forms.Cursors.No
				MousePointerID = IDC_NO
			Case System.Windows.Forms.Cursors.AppStarting
				MousePointerID = IDC_APPSTARTING
			Case System.Windows.Forms.Cursors.Help
				MousePointerID = IDC_HELP
			Case 16
				MousePointerID = IDC_WAITCD
		End Select
	End Function
	
	Public Sub RefreshMousePointer(Optional ByVal hWndFallback As Integer = 0)
		Const WM_SETCURSOR As Integer = &H20
		Const WM_NCHITTEST As Integer = &H84
		Const WM_MOUSEMOVE As Integer = &H200
		Dim P As POINTAPI
		Dim hWndCursor As Integer
		GetCursorPos(P)
		hWndCursor = GetCapture()
		If hWndCursor = 0 Then hWndCursor = WindowFromPoint(P.X, P.Y)
		If hWndCursor <> 0 Then
			If GetWindowThreadProcessId(hWndCursor, 0) <> System.Threading.Thread.CurrentThread.ManagedThreadID Then hWndCursor = hWndFallback
		Else
			hWndCursor = hWndFallback
		End If
		If hWndCursor <> 0 Then SendMessage(hWndCursor, WM_SETCURSOR, hWndCursor, MakeDWord(SendMessage(hWndCursor, WM_NCHITTEST, 0, Make_XY_lParam(P.X, P.Y)), WM_MOUSEMOVE))
	End Sub
	
	Public Function OLEFontIsEqual(ByVal Font As System.Drawing.Font, ByVal FontOther As System.Drawing.Font) As Boolean
		If Font Is Nothing Then
			If FontOther Is Nothing Then OLEFontIsEqual = True
		ElseIf FontOther Is Nothing Then 
			If Font Is Nothing Then OLEFontIsEqual = True
		Else
			'UPGRADE_ISSUE: Font property FontOther.Weight was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			'UPGRADE_ISSUE: Font property Font.Weight was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			If Font.Name = FontOther.Name And Font.SizeInPoints = FontOther.SizeInPoints And Font.GdiCharSet() = FontOther.GdiCharSet() And Font.Weight = FontOther.Weight And Font.Underline = FontOther.Underline And Font.Italic = FontOther.Italic And Font.StrikeOut = FontOther.StrikeOut Then
				OLEFontIsEqual = True
			End If
		End If
	End Function
	
	Public Function CreateGDIFontFromOLEFont(ByVal Font As System.Drawing.Font) As Integer
		If Font Is Nothing Then Exit Function
		'UPGRADE_WARNING: Arrays in structure LF may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim LF As LOGFONT
		' hFont will be cleared when the IFont reference goes out of scope or is set to nothing.
		'UPGRADE_WARNING: Couldn't resolve default property of object LF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		GetObjectAPI(Font.hFont, LenB(LF), LF)
		CreateGDIFontFromOLEFont = CreateFontIndirect(LF)
	End Function
	
	Public Function CloneOLEFont(ByVal Font As System.Drawing.Font) As System.Drawing.Font
		If Not Font Is Nothing Then Font.Clone(CloneOLEFont.Name)
	End Function
	
	Public Function GetNumberGroupDigit() As String
		GetNumberGroupDigit = Mid(FormatNumber(1000, 0,  ,  , TriState.True), 2, 1)
		If GetNumberGroupDigit = "0" Then GetNumberGroupDigit = vbNullString
	End Function
	
	Public Function GetDecimalChar() As String
		GetDecimalChar = Mid(CStr(1.1), 2, 1)
	End Function
	
	Public Function IsFormLoaded(ByVal FormName As String) As Boolean
		Dim i As Short
		For i = 0 To My.Application.OpenForms.Count - 1
			If StrComp(My.Application.OpenForms.Item(i).Name, FormName, CompareMethod.Text) = 0 Then
				IsFormLoaded = True
				Exit For
			End If
		Next i
	End Function
	
	Public Function GetWindowTitle(ByVal hWnd As Integer) As String
		Dim Buffer As String
		Buffer = New String(vbNullChar, GetWindowTextLength(hWnd) + 1)
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		GetWindowText(hWnd, StrPtr(Buffer), Len(Buffer))
		GetWindowTitle = Left(Buffer, Len(Buffer) - 1)
	End Function
	
	Public Function GetWindowClassName(ByVal hWnd As Integer) As String
		Dim Buffer As String
		Dim RetVal As Integer
		Buffer = New String(vbNullChar, 256)
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		RetVal = GetClassName(hWnd, StrPtr(Buffer), Len(Buffer))
		If RetVal <> 0 Then GetWindowClassName = Left(Buffer, RetVal)
	End Function
	
	Public Sub CenterFormToScreen(ByVal Form As System.Windows.Forms.Form, Optional ByVal RefForm As System.Windows.Forms.Form = Nothing)
		Const MONITOR_DEFAULTTOPRIMARY As Integer = &H1
		If RefForm Is Nothing Then RefForm = Form
		Dim hMonitor As Integer
		Dim MI As MONITORINFO
		Dim WndRect As RECT
		hMonitor = MonitorFromWindow(RefForm.Handle.ToInt32, MONITOR_DEFAULTTOPRIMARY)
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		MI.cbSize = LenB(MI)
		GetMonitorInfo(hMonitor, MI)
		GetWindowRect(Form.Handle.ToInt32, WndRect)
		'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Dim MDIForm As System.Windows.Forms.Form
		If TypeOf Form Is System.Windows.Forms.Form Then
			MDIForm = Form
			MDIForm.SetBounds(VB6.TwipsToPixelsX((MI.RCMonitor.Left_Renamed + (((MI.RCMonitor.Right_Renamed - MI.RCMonitor.Left_Renamed) - (WndRect.Right_Renamed - WndRect.Left_Renamed)) \ 2)) * (1440 / DPI_X())), VB6.TwipsToPixelsY((MI.RCMonitor.Top + (((MI.RCMonitor.Bottom - MI.RCMonitor.Top) - (WndRect.Bottom - WndRect.Top)) \ 2)) * (1440 / DPI_Y())), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		Else
			Form.SetBounds(VB6.TwipsToPixelsX((MI.RCMonitor.Left_Renamed + (((MI.RCMonitor.Right_Renamed - MI.RCMonitor.Left_Renamed) - (WndRect.Right_Renamed - WndRect.Left_Renamed)) \ 2)) * (1440 / DPI_X())), VB6.TwipsToPixelsY((MI.RCMonitor.Top + (((MI.RCMonitor.Bottom - MI.RCMonitor.Top) - (WndRect.Bottom - WndRect.Top)) \ 2)) * (1440 / DPI_Y())), 0, 0, Windows.Forms.BoundsSpecified.X Or Windows.Forms.BoundsSpecified.Y)
		End If
	End Sub
	
	Public Sub FlashForm(ByVal Form As System.Windows.Forms.Form)
		Const FLASHW_CAPTION As Integer = &H1
		Const FLASHW_TRAY As Integer = &H2
		Const FLASHW_TIMERNOFG As Integer = &HC
		Dim FWI As FLASHWINFO
		With FWI
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			.cbSize = LenB(FWI)
			.dwFlags = FLASHW_CAPTION Or FLASHW_TRAY Or FLASHW_TIMERNOFG
			.hWnd = Form.Handle.ToInt32
			.dwTimeout = 0 ' Default cursor blink rate
			.uCount = 0
		End With
		FlashWindowEx(FWI)
	End Sub
	
	Public Function GetFormTitleBarHeight(ByVal Form As System.Windows.Forms.Form) As Single
		Const SM_CYCAPTION As Integer = 4
		Const SM_CYMENU As Integer = 15
		Const SM_CYSIZEFRAME As Integer = 33
		Const SM_CYFIXEDFRAME As Integer = 8
		Dim CY As Integer
		CY = GetSystemMetrics(SM_CYCAPTION)
		If GetMenu(Form.Handle.ToInt32) <> 0 Then CY = CY + GetSystemMetrics(SM_CYMENU)
		Select Case Form.FormBorderStyle
			Case System.Windows.Forms.FormBorderStyle.Sizable, System.Windows.Forms.FormBorderStyle.SizableToolWindow
				CY = CY + GetSystemMetrics(SM_CYSIZEFRAME)
			Case System.Windows.Forms.FormBorderStyle.FixedSingle, System.Windows.Forms.FormBorderStyle.FixedDialog, System.Windows.Forms.FormBorderStyle.FixedToolWindow
				CY = CY + GetSystemMetrics(SM_CYFIXEDFRAME)
		End Select
		'UPGRADE_ISSUE: Form property Form.ScaleMode is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8027179A-CB3B-45C0-9863-FAA1AF983B59"'
		'UPGRADE_ISSUE: Constant vbPixels was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Form method Form.ScaleY was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		If CY > 0 Then GetFormTitleBarHeight = Form.ScaleY(CY, vbPixels, Form.ScaleMode)
	End Function
	
	Public Function GetFormNonScaleHeight(ByVal Form As System.Windows.Forms.Form) As Single
		Const SM_CYCAPTION As Integer = 4
		Const SM_CYMENU As Integer = 15
		Const SM_CYSIZEFRAME As Integer = 33
		Const SM_CYFIXEDFRAME As Integer = 8
		Dim CY As Integer
		CY = GetSystemMetrics(SM_CYCAPTION)
		If GetMenu(Form.Handle.ToInt32) <> 0 Then CY = CY + GetSystemMetrics(SM_CYMENU)
		Select Case Form.FormBorderStyle
			Case System.Windows.Forms.FormBorderStyle.Sizable, System.Windows.Forms.FormBorderStyle.SizableToolWindow
				CY = CY + (GetSystemMetrics(SM_CYSIZEFRAME) * 2)
			Case System.Windows.Forms.FormBorderStyle.FixedSingle, System.Windows.Forms.FormBorderStyle.FixedDialog, System.Windows.Forms.FormBorderStyle.FixedToolWindow
				CY = CY + (GetSystemMetrics(SM_CYFIXEDFRAME) * 2)
		End Select
		'UPGRADE_ISSUE: Form property Form.ScaleMode is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8027179A-CB3B-45C0-9863-FAA1AF983B59"'
		'UPGRADE_ISSUE: Constant vbPixels was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
		'UPGRADE_ISSUE: Form method Form.ScaleY was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
		If CY > 0 Then GetFormNonScaleHeight = Form.ScaleY(CY, vbPixels, Form.ScaleMode)
	End Function
	
	Public Sub SetWindowRedraw(ByVal hWnd As Integer, ByVal Enabled As Boolean)
		Const WM_SETREDRAW As Integer = &HB
		SendMessage(hWnd, WM_SETREDRAW, IIf(Enabled = True, 1, 0), 0)
		Const RDW_UPDATENOW As Integer = &H100
		Const RDW_INVALIDATE As Integer = &H1
		Const RDW_ERASE As Integer = &H4
		Const RDW_ALLCHILDREN As Integer = &H80
		If Enabled = True Then
			RedrawWindow(hWnd, 0, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_ERASE Or RDW_ALLCHILDREN)
		End If
	End Sub
	
	Public Function GetWindowsDir() As String
		Static Done As Boolean
		Static Value As String
		Dim Buffer As String
		If Done = False Then
			Buffer = New String(vbNullChar, MAX_PATH)
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If GetSystemWindowsDirectory(StrPtr(Buffer), MAX_PATH) <> 0 Then
				Value = Left(Buffer, InStr(Buffer, vbNullChar) - 1)
				Value = Value & IIf(Right(Value, 1) = "\", "", "\")
			End If
			Done = True
		End If
		GetWindowsDir = Value
	End Function
	
	Public Function GetSystemDir() As String
		Static Done As Boolean
		Static Value As String
		Dim Buffer As String
		If Done = False Then
			Buffer = New String(vbNullChar, MAX_PATH)
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			If GetSystemDirectory(StrPtr(Buffer), MAX_PATH) <> 0 Then
				Value = Left(Buffer, InStr(Buffer, vbNullChar) - 1)
				Value = Value & IIf(Right(Value, 1) = "\", "", "\")
			End If
			Done = True
		End If
		GetSystemDir = Value
	End Function
	
	Public Function GetShiftStateFromParam(ByVal wParam As Integer) As VB6.ShiftConstants
		Const MK_SHIFT As Integer = &H4
		Const MK_CONTROL As Integer = &H8
		If (wParam And MK_SHIFT) = MK_SHIFT Then GetShiftStateFromParam = VB6.ShiftConstants.ShiftMask
		If (wParam And MK_CONTROL) = MK_CONTROL Then GetShiftStateFromParam = GetShiftStateFromParam Or VB6.ShiftConstants.CtrlMask
		If GetKeyState(System.Windows.Forms.Keys.Menu) < 0 Then GetShiftStateFromParam = GetShiftStateFromParam Or VB6.ShiftConstants.AltMask
	End Function
	
	Public Function GetMouseStateFromParam(ByVal wParam As Integer) As VB6.MouseButtonConstants
		Const MK_LBUTTON As Integer = &H1
		Const MK_RBUTTON As Integer = &H2
		Const MK_MBUTTON As Integer = &H10
		If (wParam And MK_LBUTTON) = MK_LBUTTON Then GetMouseStateFromParam = VB6.MouseButtonConstants.LeftButton
		If (wParam And MK_RBUTTON) = MK_RBUTTON Then GetMouseStateFromParam = GetMouseStateFromParam Or VB6.MouseButtonConstants.RightButton
		If (wParam And MK_MBUTTON) = MK_MBUTTON Then GetMouseStateFromParam = GetMouseStateFromParam Or VB6.MouseButtonConstants.MiddleButton
	End Function
	
	Public Function GetShiftStateFromMsg() As VB6.ShiftConstants
		If GetKeyState(System.Windows.Forms.Keys.ShiftKey) < 0 Then GetShiftStateFromMsg = VB6.ShiftConstants.ShiftMask
		If GetKeyState(System.Windows.Forms.Keys.ControlKey) < 0 Then GetShiftStateFromMsg = GetShiftStateFromMsg Or VB6.ShiftConstants.CtrlMask
		If GetKeyState(System.Windows.Forms.Keys.Menu) < 0 Then GetShiftStateFromMsg = GetShiftStateFromMsg Or VB6.ShiftConstants.AltMask
	End Function
	
	Public Function GetMouseStateFromMsg() As VB6.MouseButtonConstants
		If GetKeyState(VB6.MouseButtonConstants.LeftButton) < 0 Then GetMouseStateFromMsg = VB6.MouseButtonConstants.LeftButton
		If GetKeyState(VB6.MouseButtonConstants.RightButton) < 0 Then GetMouseStateFromMsg = GetMouseStateFromMsg Or VB6.MouseButtonConstants.RightButton
		If GetKeyState(VB6.MouseButtonConstants.MiddleButton) < 0 Then GetMouseStateFromMsg = GetMouseStateFromMsg Or VB6.MouseButtonConstants.MiddleButton
	End Function
	
	Public Function GetShiftState() As VB6.ShiftConstants
		GetShiftState = (CShort(-VB6.ShiftConstants.ShiftMask) * CShort(KeyPressed(System.Windows.Forms.Keys.ShiftKey)))
		GetShiftState = GetShiftState Or (CShort(-VB6.ShiftConstants.CtrlMask) * CShort(KeyPressed(System.Windows.Forms.Keys.ControlKey)))
		GetShiftState = GetShiftState Or (CShort(-VB6.ShiftConstants.AltMask) * CShort(KeyPressed(System.Windows.Forms.Keys.Menu)))
	End Function
	
	Public Function GetMouseState() As VB6.MouseButtonConstants
		Const SM_SWAPBUTTON As Integer = 23
		' GetAsyncKeyState requires a mapping of physical mouse buttons to logical mouse buttons.
		GetMouseState = (CShort(-VB6.MouseButtonConstants.LeftButton) * CShort(KeyPressed(IIf(GetSystemMetrics(SM_SWAPBUTTON) = 0, VB6.MouseButtonConstants.LeftButton, VB6.MouseButtonConstants.RightButton))))
		GetMouseState = GetMouseState Or (CShort(-VB6.MouseButtonConstants.RightButton) * CShort(KeyPressed(IIf(GetSystemMetrics(SM_SWAPBUTTON) = 0, VB6.MouseButtonConstants.RightButton, VB6.MouseButtonConstants.LeftButton))))
		GetMouseState = GetMouseState Or (CShort(-VB6.MouseButtonConstants.MiddleButton) * CShort(KeyPressed(VB6.MouseButtonConstants.MiddleButton)))
	End Function
	
	Public Function KeyToggled(ByVal KeyCode As System.Windows.Forms.Keys) As Boolean
		KeyToggled = CBool(LoByte(GetKeyState(KeyCode)) = 1)
	End Function
	
	Public Function KeyPressed(ByVal KeyCode As System.Windows.Forms.Keys) As Boolean
		KeyPressed = CBool((GetAsyncKeyState(KeyCode) And &H8000) = &H8000)
	End Function
	
	Public Function InIDE(Optional ByRef B As Boolean = True) As Boolean
		If B = True Then System.Diagnostics.Debug.Assert(Not InIDE(InIDE), "") Else B = True
	End Function
	
	Public Function PtrToObj(ByVal ObjectPointer As Integer) As Object
		Dim TempObj As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object TempObj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(TempObj, ObjectPointer, 4)
		PtrToObj = TempObj
		'UPGRADE_WARNING: Couldn't resolve default property of object TempObj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(TempObj, 0, 4)
	End Function
	
	Public Function ProcPtr(ByVal Address As Integer) As Integer
		ProcPtr = Address
	End Function
	
	Public Function LoByte(ByVal Word As Short) As Byte
		LoByte = Word And &HFF
	End Function
	
	Public Function HiByte(ByVal Word As Short) As Byte
		HiByte = (Word And &HFF00) \ &H100
	End Function
	
	Public Function MakeWord(ByVal LoByte As Byte, ByVal HiByte As Byte) As Short
		If (HiByte And &H80) <> 0 Then
			MakeWord = ((HiByte * &H100) Or LoByte) Or &HFFFF0000
		Else
			MakeWord = (HiByte * &H100) Or LoByte
		End If
	End Function
	
	Public Function LoWord(ByVal DWord As Integer) As Short
		If DWord And &H8000 Then
			LoWord = DWord Or &HFFFF0000
		Else
			LoWord = DWord And &HFFFF
		End If
	End Function
	
	Public Function HiWord(ByVal DWord As Integer) As Short
		HiWord = (DWord And &HFFFF0000) \ &H10000
	End Function
	
	Public Function MakeDWord(ByVal LoWord As Short, ByVal HiWord As Short) As Integer
		MakeDWord = (CInt(HiWord) * &H10000) Or (LoWord And &HFFFF)
	End Function
	
	Public Function Get_X_lParam(ByVal lParam As Integer) As Integer
		Get_X_lParam = lParam And &H7FFF
		If lParam And &H8000 Then Get_X_lParam = Get_X_lParam Or &HFFFF8000
	End Function
	
	Public Function Get_Y_lParam(ByVal lParam As Integer) As Integer
		Get_Y_lParam = (lParam And &H7FFF0000) \ &H10000
		If lParam And &H80000000 Then Get_Y_lParam = Get_Y_lParam Or &HFFFF8000
	End Function
	
	Public Function Make_XY_lParam(ByVal X As Integer, ByVal Y As Integer) As Integer
		Make_XY_lParam = LoWord(X) Or (&H10000 * LoWord(Y))
	End Function
	
	Public Function UTF32CodePoint_To_UTF16(ByVal CodePoint As Integer) As String
		Dim HW, LW As Short
		If CodePoint >= &HFFFF8000 And CodePoint <= &H10FFFF Then
			If CodePoint < &H10000 Then
				HW = 0
				LW = CUIntToInt(CodePoint And &HFFFF)
			Else
				CodePoint = CodePoint - &H10000
				HW = (CodePoint \ &H400) + &HD800
				LW = (CodePoint Mod &H400) + &HDC00
			End If
			If HW = 0 Then UTF32CodePoint_To_UTF16 = ChrW(LW) Else UTF32CodePoint_To_UTF16 = ChrW(HW) & ChrW(LW)
		End If
	End Function
	
	Public Function UTF16_To_UTF8(ByRef Source As String) As Byte()
		Const CP_UTF8 As Integer = 65001
		Dim Pointer, Length, Size As Integer
		Length = Len(Source)
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		Pointer = StrPtr(Source)
		Size = WideCharToMultiByte(CP_UTF8, 0, Pointer, Length, 0, 0, 0, 0)
		Dim Buffer() As Byte
		If Size > 0 Then
			ReDim Buffer(Size - 1)
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			WideCharToMultiByte(CP_UTF8, 0, Pointer, Length, VarPtr(Buffer(0)), Size, 0, 0)
			UTF16_To_UTF8 = VB6.CopyArray(Buffer)
		End If
	End Function
	
	Public Function UTF8_To_UTF16(ByRef Source() As Byte) As String
		If (0 / 1) + CShort(Source) = 0 Then Exit Function
		Const CP_UTF8 As Integer = 65001
		Dim Pointer, Size, Length As Integer
		Size = UBound(Source) - LBound(Source) + 1
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		Pointer = VarPtr(Source(LBound(Source)))
		Length = MultiByteToWideChar(CP_UTF8, 0, Pointer, Size, 0, 0)
		If Length > 0 Then
			UTF8_To_UTF16 = Space(Length)
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			MultiByteToWideChar(CP_UTF8, 0, Pointer, Size, StrPtr(UTF8_To_UTF16), Length)
		End If
	End Function
	
	Public Function StrToVar(ByVal Text As String) As Object
		Dim B() As Byte
		If Text = vbNullString Then
			'UPGRADE_WARNING: Couldn't resolve default property of object StrToVar. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			StrToVar = Nothing
		Else
			'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetBytes() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
			B = System.Text.UnicodeEncoding.Unicode.GetBytes(Text)
			StrToVar = VB6.CopyArray(B)
		End If
	End Function
	
	Public Function VarToStr(ByVal Bytes As Object) As String
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Dim B() As Byte
		If IsNothing(Bytes) Then
			VarToStr = vbNullString
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Bytes. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			B = Bytes
			'UPGRADE_TODO: Code was upgraded to use System.Text.UnicodeEncoding.Unicode.GetString() which may not have the same behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93DD716C-10E3-41BE-A4A8-3BA40157905B"'
			VarToStr = System.Text.UnicodeEncoding.Unicode.GetString(B)
		End If
	End Function
	
	Public Function UnsignedAdd(ByVal Start As Integer, ByVal Incr As Integer) As Integer
		UnsignedAdd = (CShort(Start Xor &H80000000) + Incr) Xor &H80000000
	End Function
	
	Public Function UnsignedSub(ByVal Start As Integer, ByVal Decr As Integer) As Integer
		UnsignedSub = (CShort(Start And &H7FFFFFFF) - CShort(Decr And &H7FFFFFFF)) Xor ((Start Xor Decr) And &H80000000)
	End Function
	
	Public Function CUIntToInt(ByVal Value As Integer) As Short
		Const OFFSET_2 As Integer = 65536
		Const MAXINT_2 As Short = 32767
		If Value < 0 Or Value >= OFFSET_2 Then Err.Raise(6)
		If Value <= MAXINT_2 Then
			CUIntToInt = Value
		Else
			CUIntToInt = Value - OFFSET_2
		End If
	End Function
	
	Public Function CIntToUInt(ByVal Value As Short) As Integer
		Const OFFSET_2 As Integer = 65536
		If Value < 0 Then
			CIntToUInt = Value + OFFSET_2
		Else
			CIntToUInt = Value
		End If
	End Function
	
	Public Function CULngToLng(ByVal Value As Double) As Integer
		Const OFFSET_4 As Double = 4294967296#
		Const MAXINT_4 As Integer = 2147483647
		If Value < 0 Or Value >= OFFSET_4 Then Err.Raise(6)
		If Value <= MAXINT_4 Then
			CULngToLng = Value
		Else
			CULngToLng = Value - OFFSET_4
		End If
	End Function
	
	Public Function CLngToULng(ByVal Value As Integer) As Double
		Const OFFSET_4 As Double = 4294967296#
		If Value < 0 Then
			CLngToULng = Value + OFFSET_4
		Else
			CLngToULng = Value
		End If
	End Function
	
	Public Function DPI_X() As Integer
		Const LOGPIXELSX As Integer = 88
		Dim hDCScreen As Integer
		hDCScreen = GetDC(0)
		If hDCScreen <> 0 Then
			DPI_X = GetDeviceCaps(hDCScreen, LOGPIXELSX)
			ReleaseDC(0, hDCScreen)
		End If
	End Function
	
	Public Function DPI_Y() As Integer
		Const LOGPIXELSY As Integer = 90
		Dim hDCScreen As Integer
		hDCScreen = GetDC(0)
		If hDCScreen <> 0 Then
			DPI_Y = GetDeviceCaps(hDCScreen, LOGPIXELSY)
			ReleaseDC(0, hDCScreen)
		End If
	End Function
	
	Public Function DPICorrectionFactor() As Single
		Static Done As Boolean
		Static Value As Single
		If Done = False Then
			Value = ((96 / DPI_X()) * 15) / VB6.TwipsPerPixelX
			Done = True
		End If
		' Returns exactly 1 when no corrections are required.
		DPICorrectionFactor = Value
	End Function
	
	Public Function CHimetricToPixel_X(ByVal Width As Integer) As Integer
		Const HIMETRIC_PER_INCH As Integer = 2540
		CHimetricToPixel_X = (Width * DPI_X()) / HIMETRIC_PER_INCH
	End Function
	
	Public Function CHimetricToPixel_Y(ByVal Height As Integer) As Integer
		Const HIMETRIC_PER_INCH As Integer = 2540
		CHimetricToPixel_Y = (Height * DPI_Y()) / HIMETRIC_PER_INCH
	End Function
	
	Public Function PixelsPerDIP_X() As Single
		Static Done As Boolean
		Static Value As Single
		If Done = False Then
			Value = (DPI_X() / 96)
			Done = True
		End If
		PixelsPerDIP_X = Value
	End Function
	
	Public Function PixelsPerDIP_Y() As Single
		Static Done As Boolean
		Static Value As Single
		If Done = False Then
			Value = (DPI_Y() / 96)
			Done = True
		End If
		PixelsPerDIP_Y = Value
	End Function
	
	Public Function WinColor(ByVal Color As Integer, Optional ByVal hPal As Integer = 0) As Integer
		If OleTranslateColor(Color, hPal, WinColor) <> 0 Then WinColor = -1
	End Function
	
	Public Function PictureFromByteStream(ByRef ByteStream As Object) As System.Drawing.Image
		Const GMEM_MOVEABLE As Integer = &H2
		'UPGRADE_WARNING: Arrays in structure IID may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim IID As CLSID
		Dim Stream As stdole.IUnknown
		Dim NewPicture As System.Drawing.Image
		Dim B() As Byte
		Dim ByteCount As Integer
		Dim hMem, lpMem As Integer
		With IID
			.Data1 = &H7BF80980
			.Data2 = &HBF32
			.Data3 = &H101A
			.Data4(0) = &H8B
			.Data4(1) = &HBB
			.Data4(3) = &HAA
			.Data4(5) = &H30
			.Data4(6) = &HC
			.Data4(7) = &HAB
		End With
		'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If VarType(ByteStream) = (VariantType.Array + VariantType.Byte) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object ByteStream. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			B = ByteStream
			ByteCount = (UBound(B) - LBound(B)) + 1
			hMem = GlobalAlloc(GMEM_MOVEABLE, ByteCount)
			If hMem <> 0 Then
				lpMem = GlobalLock(hMem)
				If lpMem <> 0 Then
					CopyMemory(lpMem, B(LBound(B)), ByteCount)
					GlobalUnlock(hMem)
					If CreateStreamOnHGlobal(hMem, 1, Stream) = 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object IID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If OleLoadPicture(Stream, ByteCount, 0, IID, NewPicture) = 0 Then PictureFromByteStream = NewPicture
					End If
				End If
			End If
		End If
	End Function
	
	Public Function PictureFromPath(ByVal PathName As String) As System.Drawing.Image
		'UPGRADE_WARNING: Arrays in structure IID may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim IID As CLSID
		Dim NewPicture As System.Drawing.Image
		With IID
			.Data1 = &H7BF80980
			.Data2 = &HBF32
			.Data3 = &H101A
			.Data4(0) = &H8B
			.Data4(1) = &HBB
			.Data4(3) = &HAA
			.Data4(5) = &H30
			.Data4(6) = &HC
			.Data4(7) = &HAB
		End With
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If OleLoadPicturePath(StrPtr(PathName), 0, 0, System.Drawing.ColorTranslator.FromOle(0), IID, NewPicture) = 0 Then PictureFromPath = NewPicture
	End Function
	
	'UPGRADE_ISSUE: VBRUN.PictureTypeConstants type was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
	Public Function PictureFromHandle(ByVal Handle As Integer, ByVal PicType As Object) As System.Drawing.Image
		If Handle = 0 Then Exit Function
		'UPGRADE_WARNING: Arrays in structure IID may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim PICD As PICTDESC
		Dim IID As CLSID
		Dim NewPicture As System.Drawing.Image
		With PICD
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			.cbSizeOfStruct = LenB(PICD)
			.PicType = PicType
			.hImage = Handle
		End With
		With IID
			.Data1 = &H7BF80980
			.Data2 = &HBF32
			.Data3 = &H101A
			.Data4(0) = &H8B
			.Data4(1) = &HBB
			.Data4(3) = &HAA
			.Data4(5) = &H30
			.Data4(6) = &HC
			.Data4(7) = &HAB
		End With
		'UPGRADE_WARNING: Couldn't resolve default property of object IID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If OleCreatePictureIndirect(PICD, IID, 1, NewPicture) = 0 Then PictureFromHandle = NewPicture
	End Function
	
	Public Function BitmapHandleFromPicture(ByVal Picture As System.Drawing.Image, Optional ByVal BackColor As System.Drawing.Color = Nothing) As Integer
		If Picture Is Nothing Then Exit Function
		Dim hBmp, hDCScreen, hDC, hBmpOld As Integer
		Dim CY, CX, Brush As Integer
		Const DI_NORMAL As Integer = &H3
		Dim RC As RECT
		With Picture
			'UPGRADE_ISSUE: Picture property Picture.Handle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			If .Handle <> 0 Then
				'UPGRADE_ISSUE: Picture property Picture.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				CX = CHimetricToPixel_X(.Width)
				'UPGRADE_ISSUE: Picture property Picture.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				CY = CHimetricToPixel_Y(.Height)
				Brush = CreateSolidBrush(WinColor(System.Drawing.ColorTranslator.ToOle(BackColor)))
				hDCScreen = GetDC(0)
				If hDCScreen <> 0 Then
					hDC = CreateCompatibleDC(hDCScreen)
					If hDC <> 0 Then
						hBmp = CreateCompatibleBitmap(hDCScreen, CX, CY)
						If hBmp <> 0 Then
							hBmpOld = SelectObject(hDC, hBmp)
							'UPGRADE_ISSUE: Constant vbPicTypeIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
							'UPGRADE_ISSUE: Picture property Picture.Type was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							If .Type = vbPicTypeIcon Then
								'UPGRADE_ISSUE: Picture property Picture.Handle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
								DrawIconEx(hDC, 0, 0, .Handle, CX, CY, 0, Brush, DI_NORMAL)
							Else
								RC.Right_Renamed = CX
								RC.Bottom = CY
								FillRect(hDC, RC, Brush)
								'UPGRADE_ISSUE: Picture property Picture.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
								'UPGRADE_ISSUE: Picture property Picture.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
								'UPGRADE_ISSUE: Picture method Picture.Render was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
								.Render(hDC Or 0, 0, 0, CX Or 0, CY Or 0, 0, .Height, .Width, -.Height, 0)
							End If
							SelectObject(hDC, hBmpOld)
							BitmapHandleFromPicture = hBmp
						End If
						DeleteDC(hDC)
					End If
					ReleaseDC(0, hDCScreen)
				End If
				DeleteObject(Brush)
			End If
		End With
	End Function
	
	Public Sub RenderPicture(ByVal Picture As System.Drawing.Image, ByVal hDC As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal CX As Integer = 0, Optional ByVal CY As Integer = 0, Optional ByRef RenderFlag As Short = 0)
		' RenderFlag is passed as a optional parameter ByRef.
		' It is ignored for icons and metafiles.
		' 0 = render method unknown, determine it and update parameter
		' 1 = StdPicture.Render
		' 2 = GdiAlphaBlend
		If Picture Is Nothing Then Exit Sub
		Const DI_NORMAL As Integer = &H3
		Dim HasAlpha As Boolean
		Const PICTURE_TRANSPARENT As Integer = &H2
		Dim Bmp As BITMAP
		Dim j, i, Pos As Integer
		Dim hDCBmp, hBmpOld As Integer
		Dim SA1D As SAFEARRAY1D
		Dim B() As Byte ' Exclude GIF
		With Picture
			'UPGRADE_ISSUE: Picture property Picture.Handle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
			If .Handle <> 0 Then
				'UPGRADE_ISSUE: Picture property Picture.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				If CX = 0 Then CX = CHimetricToPixel_X(.Width)
				'UPGRADE_ISSUE: Picture property Picture.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				If CY = 0 Then CY = CHimetricToPixel_Y(.Height)
				'UPGRADE_ISSUE: Constant vbPicTypeIcon was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
				'UPGRADE_ISSUE: Picture property Picture.Type was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
				If .Type = vbPicTypeIcon Then
					'UPGRADE_ISSUE: Picture property Picture.Handle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					DrawIconEx(hDC, X, Y, .Handle, CX, CY, 0, 0, DI_NORMAL)
				Else
					'UPGRADE_ISSUE: Constant vbPicTypeBitmap was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
					'UPGRADE_ISSUE: Picture property Picture.Type was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
					If .Type = vbPicTypeBitmap Then
						If RenderFlag = 0 Then
							If (.Attributes And PICTURE_TRANSPARENT) = 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object Bmp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
								'UPGRADE_ISSUE: Picture property Picture.Handle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
								GetObjectAPI(.Handle, LenB(Bmp), Bmp)
								If Bmp.BMBitsPixel = 32 And Bmp.BMBits <> 0 Then
									With SA1D
										.cDims = 1
										.fFeatures = 0
										.cbElements = 1
										.cLocks = 0
										.pvData = Bmp.BMBits
										.Bounds.lLbound = 0
										.Bounds.cElements = Bmp.BMWidthBytes * Bmp.BMHeight
									End With
									'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
									CopyMemory(ArrPtr(B), VarPtr(SA1D), 4)
									For i = 0 To (System.Math.Abs(Bmp.BMHeight) - 1)
										Pos = i * Bmp.BMWidthBytes
										For j = (Pos + 3) To (Pos + Bmp.BMWidthBytes - 1) Step 4
											If HasAlpha = False Then HasAlpha = (B(j) > 0)
											If HasAlpha = True Then
												If B(j - 1) > B(j) Then
													HasAlpha = False
													i = System.Math.Abs(Bmp.BMHeight) - 1
													Exit For
												ElseIf B(j - 2) > B(j) Then 
													HasAlpha = False
													i = System.Math.Abs(Bmp.BMHeight) - 1
													Exit For
												ElseIf B(j - 3) > B(j) Then 
													HasAlpha = False
													i = System.Math.Abs(Bmp.BMHeight) - 1
													Exit For
												End If
											End If
										Next j
									Next i
									CopyMemory(ArrPtr(B), 0, 4)
								End If
							End If
							If HasAlpha = False Then RenderFlag = 1 Else RenderFlag = 2
						ElseIf RenderFlag = 2 Then 
							HasAlpha = True
						End If
					End If
					If HasAlpha = False Then
						'UPGRADE_ISSUE: Picture property Picture.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_ISSUE: Picture property Picture.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						'UPGRADE_ISSUE: Picture method Picture.Render was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
						.Render(hDC Or 0, X Or 0, Y Or 0, CX Or 0, CY Or 0, 0, .Height, .Width, -.Height, 0)
					Else
						hDCBmp = CreateCompatibleDC(0)
						If hDCBmp <> 0 Then
							'UPGRADE_ISSUE: Picture property Picture.Handle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							hBmpOld = SelectObject(hDCBmp, .Handle)
							'UPGRADE_ISSUE: Picture property Picture.Height was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							'UPGRADE_ISSUE: Picture property Picture.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
							GdiAlphaBlend(hDC, X, Y, CX, CY, hDCBmp, 0, 0, CHimetricToPixel_X(.Width), CHimetricToPixel_Y(.Height), &H1FF0000)
							SelectObject(hDCBmp, hBmpOld)
							DeleteDC(hDCBmp)
						End If
					End If
				End If
			End If
		End With
	End Sub
End Module