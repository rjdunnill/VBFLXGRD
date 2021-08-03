Option Strict Off
Option Explicit On
Module VTableHandle
	
	' Required:
	
	' OLEGuids.tlb (in IDE only)
	
#If False Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression False did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Private VTableInterfaceInPlaceActiveObject, VTableInterfaceControl, VTableInterfacePerPropertyBrowsing
#End If
	Public Enum VTableInterfaceConstants
		VTableInterfaceInPlaceActiveObject = 1
		VTableInterfaceControl = 2
		VTableInterfacePerPropertyBrowsing = 3
	End Enum
	Private Structure VTableIPAODataStruct
		Dim VTable As Integer
		Dim RefCount As Integer
		Dim OriginalIOleIPAO As OLEGuids.IOleInPlaceActiveObject
		Dim IOleIPAO As OLEGuids.IOleInPlaceActiveObjectVB
	End Structure
	Private Structure VTableEnumVARIANTDataStruct
		Dim VTable As Integer
		Dim RefCount As Integer
		Dim Enumerable As Object
		Dim Index As Integer
		Dim Count As Integer
	End Structure
	Public Const CTRLINFO_EATS_RETURN As Integer = 1
	Public Const CTRLINFO_EATS_ESCAPE As Integer = 2
	Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Integer)
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef Destination As Any, ByRef Source As Any, ByVal Length As Integer)
	Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Integer)
	Private Declare Function CoTaskMemAlloc Lib "ole32" (ByVal cBytes As Integer) As Integer
	Private Declare Function SysAllocString Lib "oleaut32" (ByVal lpString As Integer) As Integer
	Private Declare Function DispCallFunc Lib "oleaut32" (ByVal lpvInstance As Integer, ByVal oVft As Integer, ByVal CallConv As Integer, ByVal vtReturn As Short, ByVal cActuals As Integer, ByVal prgvt As Integer, ByVal prgpvarg As Integer, ByRef pvargResult As Object) As Integer
	Private Declare Function VariantCopyToPtr Lib "oleaut32"  Alias "VariantCopy"(ByVal pvargDest As Integer, ByRef pvargSrc As Object) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Integer, ByRef pCLSID As Any) As Integer
	Private Const CC_STDCALL As Integer = 4
	Private Const E_OUTOFMEMORY As Integer = &H8007000E
	Private Const E_INVALIDARG As Integer = &H80070057
	Private Const E_NOTIMPL As Integer = &H80004001
	Private Const E_NOINTERFACE As Integer = &H80004002
	Private Const E_POINTER As Integer = &H80004003
	Private Const S_FALSE As Integer = &H1
	Private Const S_OK As Integer = &H0
	Private VTableIPAO(9) As Integer
	Private VTableIPAOData As VTableIPAODataStruct
	Private VTableControl(6) As Integer
	Private OriginalVTableControl As Integer
	Private VTablePPB(6) As Integer
	Private OriginalVTablePPB As Integer
	Private StringsOutArray() As String
	Private CookiesOutArray() As Integer
	Private VTableEnumVARIANT(6) As Integer
	
	Public Function SetVTableHandling(ByVal This As Object, ByVal OLEInterface As VTableInterfaceConstants) As Boolean
		Select Case OLEInterface
			Case VTableInterfaceConstants.VTableInterfaceInPlaceActiveObject
				If VTableHandlingSupported(This, VTableInterfaceConstants.VTableInterfaceInPlaceActiveObject) = True Then
					VTableIPAOData.RefCount = VTableIPAOData.RefCount + 1
					SetVTableHandling = True
				End If
			Case VTableInterfaceConstants.VTableInterfaceControl
				If VTableHandlingSupported(This, VTableInterfaceConstants.VTableInterfaceControl) = True Then
					'UPGRADE_WARNING: Couldn't resolve default property of object This. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Call ReplaceIOleControl(This)
					SetVTableHandling = True
				End If
			Case VTableInterfaceConstants.VTableInterfacePerPropertyBrowsing
				If VTableHandlingSupported(This, VTableInterfaceConstants.VTableInterfacePerPropertyBrowsing) = True Then
					'UPGRADE_WARNING: Couldn't resolve default property of object This. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Call ReplaceIPPB(This)
					SetVTableHandling = True
				End If
		End Select
	End Function
	
	Public Function RemoveVTableHandling(ByVal This As Object, ByVal OLEInterface As VTableInterfaceConstants) As Boolean
		Select Case OLEInterface
			Case VTableInterfaceConstants.VTableInterfaceInPlaceActiveObject
				If VTableHandlingSupported(This, VTableInterfaceConstants.VTableInterfaceInPlaceActiveObject) = True Then
					VTableIPAOData.RefCount = VTableIPAOData.RefCount - 1
					RemoveVTableHandling = True
				End If
			Case VTableInterfaceConstants.VTableInterfaceControl
				If VTableHandlingSupported(This, VTableInterfaceConstants.VTableInterfaceControl) = True Then
					'UPGRADE_WARNING: Couldn't resolve default property of object This. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Call RestoreIOleControl(This)
					RemoveVTableHandling = True
				End If
			Case VTableInterfaceConstants.VTableInterfacePerPropertyBrowsing
				If VTableHandlingSupported(This, VTableInterfaceConstants.VTableInterfacePerPropertyBrowsing) = True Then
					'UPGRADE_WARNING: Couldn't resolve default property of object This. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Call RestoreIPPB(This)
					RemoveVTableHandling = True
				End If
		End Select
	End Function
	
	Private Function VTableHandlingSupported(ByRef This As Object, ByVal OLEInterface As VTableInterfaceConstants) As Boolean
		On Error GoTo CATCH_EXCEPTION
		Dim ShadowIOleIPAO As OLEGuids.IOleInPlaceActiveObject
		Dim ShadowIOleInPlaceActiveObjectVB As OLEGuids.IOleInPlaceActiveObjectVB
		Dim ShadowIOleControl As OLEGuids.IOleControl
		Dim ShadowIOleControlVB As OLEGuids.IOleControlVB
		Dim ShadowIPPB As OLEGuids.IPerPropertyBrowsing
		Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB
		Select Case OLEInterface
			Case VTableInterfaceConstants.VTableInterfaceInPlaceActiveObject
				ShadowIOleIPAO = This
				ShadowIOleInPlaceActiveObjectVB = This
				VTableHandlingSupported = Not CBool(ShadowIOleIPAO Is Nothing Or ShadowIOleInPlaceActiveObjectVB Is Nothing)
			Case VTableInterfaceConstants.VTableInterfaceControl
				ShadowIOleControl = This
				ShadowIOleControlVB = This
				VTableHandlingSupported = Not CBool(ShadowIOleControl Is Nothing Or ShadowIOleControlVB Is Nothing)
			Case VTableInterfaceConstants.VTableInterfacePerPropertyBrowsing
				ShadowIPPB = This
				ShadowIPerPropertyBrowsingVB = This
				VTableHandlingSupported = Not CBool(ShadowIPPB Is Nothing Or ShadowIPerPropertyBrowsingVB Is Nothing)
		End Select
CATCH_EXCEPTION: 
	End Function
	
	'UPGRADE_WARNING: ParamArray ArgList was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Public Function VTableCall(ByVal RetType As VariantType, ByVal InterfacePointer As Integer, ByVal Entry As Integer, ParamArray ByVal ArgList() As Object) As Object
		System.Diagnostics.Debug.Assert(Not (Entry < 1 Or InterfacePointer = 0), "")
		Dim VarArgList As Object
		Dim HResult As Integer
		'UPGRADE_WARNING: Couldn't resolve default property of object VarArgList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		VarArgList = ArgList
		Dim i As Integer
		Dim ArrVarType() As Short
		Dim ArrVarPtr() As Integer
		If UBound(VarArgList) > -1 Then
			'UPGRADE_WARNING: Lower bound of array ArrVarType was changed from LBound(VarArgList) to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim ArrVarType(UBound(VarArgList))
			'UPGRADE_WARNING: Lower bound of array ArrVarPtr was changed from LBound(VarArgList) to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim ArrVarPtr(UBound(VarArgList))
			For i = LBound(VarArgList) To UBound(VarArgList)
				'UPGRADE_WARNING: VarType has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
				ArrVarType(i) = VarType(VarArgList(i))
				'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				ArrVarPtr(i) = VarPtr(VarArgList(i))
			Next i
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			HResult = DispCallFunc(InterfacePointer, (Entry - 1) * 4, CC_STDCALL, RetType, i, VarPtr(ArrVarType(0)), VarPtr(ArrVarPtr(0)), VTableCall)
		Else
			HResult = DispCallFunc(InterfacePointer, (Entry - 1) * 4, CC_STDCALL, RetType, 0, 0, 0, VTableCall)
		End If
		SetLastError(HResult) ' S_OK will clear the last error code, if any.
	End Function
	
	Public Function VTableInterfaceSupported(ByVal This As OLEGuids.IUnknownUnrestricted, ByVal IIDString As String) As Boolean
		System.Diagnostics.Debug.Assert(Not (This Is Nothing), "")
		'UPGRADE_WARNING: Arrays in structure IID may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim HResult, ObjectPointer As Integer
		Dim IID As OLEGuids.OLECLSID
		'UPGRADE_WARNING: Couldn't resolve default property of object IID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		CLSIDFromString(StrPtr(IIDString), IID)
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		HResult = This.QueryInterface(VarPtr(IID), ObjectPointer)
		Dim IUnk As OLEGuids.IUnknownUnrestricted
		If ObjectPointer <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object IUnk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(IUnk, ObjectPointer, 4)
			IUnk.Release()
			'UPGRADE_WARNING: Couldn't resolve default property of object IUnk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(IUnk, 0, 4)
		End If
		VTableInterfaceSupported = CBool(HResult = S_OK)
	End Function
	
	Public Sub SyncObjectRectsToContainer(ByVal This As Object)
		On Error GoTo CATCH_EXCEPTION
		Dim PropOleObject As OLEGuids.IOleObject
		Dim PropOleInPlaceObject As OLEGuids.IOleInPlaceObject
		Dim PropOleInPlaceSite As OLEGuids.IOleInPlaceSite
		Dim PosRect As OLEGuids.OLERECT
		Dim ClipRect As OLEGuids.OLERECT
		Dim FrameInfo As OLEGuids.OLEINPLACEFRAMEINFO
		PropOleObject = This
		PropOleInPlaceObject = This
		PropOleInPlaceSite = PropOleObject.GetClientSite
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		PropOleInPlaceSite.GetWindowContext(Nothing, Nothing, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo))
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		PropOleInPlaceObject.SetObjectRects(VarPtr(PosRect), VarPtr(ClipRect))
CATCH_EXCEPTION: 
	End Sub
	
	Public Sub ActivateIPAO(ByVal This As Object)
		On Error GoTo CATCH_EXCEPTION
		Dim PropOleObject As OLEGuids.IOleObject
		Dim PropOleInPlaceSite As OLEGuids.IOleInPlaceSite
		Dim PropOleInPlaceFrame As OLEGuids.IOleInPlaceFrame
		Dim PropOleInPlaceUIWindow As OLEGuids.IOleInPlaceUIWindow
		Dim PropOleInPlaceActiveObject As OLEGuids.IOleInPlaceActiveObject
		Dim PosRect As OLEGuids.OLERECT
		Dim ClipRect As OLEGuids.OLERECT
		Dim FrameInfo As OLEGuids.OLEINPLACEFRAMEINFO
		PropOleObject = This
		If VTableIPAOData.RefCount > 0 Then
			With VTableIPAOData
				.VTable = GetVTableIPAO()
				.OriginalIOleIPAO = This
				.IOleIPAO = This
			End With
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			CopyMemory(VarPtr(PropOleInPlaceActiveObject), VarPtr(VTableIPAOData), 4)
			'UPGRADE_WARNING: Couldn't resolve default property of object PropOleInPlaceActiveObject.AddRef. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			PropOleInPlaceActiveObject.AddRef()
		Else
			PropOleInPlaceActiveObject = This
		End If
		PropOleInPlaceSite = PropOleObject.GetClientSite
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		PropOleInPlaceSite.GetWindowContext(PropOleInPlaceFrame, PropOleInPlaceUIWindow, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo))
		'UPGRADE_WARNING: Couldn't resolve default property of object PropOleInPlaceFrame.SetActiveObject. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PropOleInPlaceFrame.SetActiveObject(PropOleInPlaceActiveObject, vbNullString)
		If Not PropOleInPlaceUIWindow Is Nothing Then PropOleInPlaceUIWindow.SetActiveObject(PropOleInPlaceActiveObject, vbNullString)
CATCH_EXCEPTION: 
	End Sub
	
	Public Sub DeActivateIPAO()
		On Error GoTo CATCH_EXCEPTION
		If VTableIPAOData.OriginalIOleIPAO Is Nothing Then Exit Sub
		Dim PropOleObject As OLEGuids.IOleObject
		Dim PropOleInPlaceSite As OLEGuids.IOleInPlaceSite
		Dim PropOleInPlaceFrame As OLEGuids.IOleInPlaceFrame
		Dim PropOleInPlaceUIWindow As OLEGuids.IOleInPlaceUIWindow
		Dim PosRect As OLEGuids.OLERECT
		Dim ClipRect As OLEGuids.OLERECT
		Dim FrameInfo As OLEGuids.OLEINPLACEFRAMEINFO
		PropOleObject = VTableIPAOData.OriginalIOleIPAO
		PropOleInPlaceSite = PropOleObject.GetClientSite
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		PropOleInPlaceSite.GetWindowContext(PropOleInPlaceFrame, PropOleInPlaceUIWindow, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo))
		'UPGRADE_WARNING: Couldn't resolve default property of object PropOleInPlaceFrame.SetActiveObject. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PropOleInPlaceFrame.SetActiveObject(Nothing, vbNullString)
		If Not PropOleInPlaceUIWindow Is Nothing Then PropOleInPlaceUIWindow.SetActiveObject(Nothing, vbNullString)
CATCH_EXCEPTION: 
		'UPGRADE_NOTE: Object VTableIPAOData.OriginalIOleIPAO may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		VTableIPAOData.OriginalIOleIPAO = Nothing
		'UPGRADE_NOTE: Object VTableIPAOData.IOleIPAO may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		VTableIPAOData.IOleIPAO = Nothing
	End Sub
	
	Private Function GetVTableIPAO() As Integer
		If VTableIPAO(0) = 0 Then
			'UPGRADE_WARNING: Add a delegate for AddressOf IOleIPAO_QueryInterface Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableIPAO(0) = ProcPtr(AddressOf IOleIPAO_QueryInterface)
			'UPGRADE_WARNING: Add a delegate for AddressOf IOleIPAO_AddRef Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableIPAO(1) = ProcPtr(AddressOf IOleIPAO_AddRef)
			'UPGRADE_WARNING: Add a delegate for AddressOf IOleIPAO_Release Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableIPAO(2) = ProcPtr(AddressOf IOleIPAO_Release)
			'UPGRADE_WARNING: Add a delegate for AddressOf IOleIPAO_GetWindow Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableIPAO(3) = ProcPtr(AddressOf IOleIPAO_GetWindow)
			'UPGRADE_WARNING: Add a delegate for AddressOf IOleIPAO_ContextSensitiveHelp Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableIPAO(4) = ProcPtr(AddressOf IOleIPAO_ContextSensitiveHelp)
			'UPGRADE_WARNING: Add a delegate for AddressOf IOleIPAO_TranslateAccelerator Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableIPAO(5) = ProcPtr(AddressOf IOleIPAO_TranslateAccelerator)
			'UPGRADE_WARNING: Add a delegate for AddressOf IOleIPAO_OnFrameWindowActivate Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableIPAO(6) = ProcPtr(AddressOf IOleIPAO_OnFrameWindowActivate)
			'UPGRADE_WARNING: Add a delegate for AddressOf IOleIPAO_OnDocWindowActivate Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableIPAO(7) = ProcPtr(AddressOf IOleIPAO_OnDocWindowActivate)
			'UPGRADE_WARNING: Add a delegate for AddressOf IOleIPAO_ResizeBorder Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableIPAO(8) = ProcPtr(AddressOf IOleIPAO_ResizeBorder)
			'UPGRADE_WARNING: Add a delegate for AddressOf IOleIPAO_EnableModeless Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableIPAO(9) = ProcPtr(AddressOf IOleIPAO_EnableModeless)
		End If
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		GetVTableIPAO = VarPtr(VTableIPAO(0))
	End Function
	
	Private Function IOleIPAO_QueryInterface(ByRef This As VTableIPAODataStruct, ByRef IID As OLEGuids.OLECLSID, ByRef pvObj As Integer) As Integer
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If VarPtr(pvObj) = 0 Then
			IOleIPAO_QueryInterface = E_POINTER
			Exit Function
		End If
		' IID_IOleInPlaceActiveObject = {00000117-0000-0000-C000-000000000046}
		If IID.Data1 = &H117 And IID.Data2 = &H0 And IID.Data3 = &H0 Then
			If IID.Data4(0) = &HC0 And IID.Data4(1) = &H0 And IID.Data4(2) = &H0 And IID.Data4(3) = &H0 And IID.Data4(4) = &H0 And IID.Data4(5) = &H0 And IID.Data4(6) = &H0 And IID.Data4(7) = &H46 Then
				'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				pvObj = VarPtr(This)
				IOleIPAO_AddRef(This)
				IOleIPAO_QueryInterface = S_OK
			Else
				'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object This.OriginalIOleIPAO.QueryInterface. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				IOleIPAO_QueryInterface = This.OriginalIOleIPAO.QueryInterface(VarPtr(IID), pvObj)
			End If
		Else
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object This.OriginalIOleIPAO.QueryInterface. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			IOleIPAO_QueryInterface = This.OriginalIOleIPAO.QueryInterface(VarPtr(IID), pvObj)
		End If
	End Function
	
	Private Function IOleIPAO_AddRef(ByRef This As VTableIPAODataStruct) As Integer
		'UPGRADE_WARNING: Couldn't resolve default property of object This.OriginalIOleIPAO.AddRef. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		IOleIPAO_AddRef = This.OriginalIOleIPAO.AddRef
	End Function
	
	Private Function IOleIPAO_Release(ByRef This As VTableIPAODataStruct) As Integer
		'UPGRADE_WARNING: Couldn't resolve default property of object This.OriginalIOleIPAO.Release. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		IOleIPAO_Release = This.OriginalIOleIPAO.Release
	End Function
	
	Private Function IOleIPAO_GetWindow(ByRef This As VTableIPAODataStruct, ByRef hWnd As Integer) As Integer
		IOleIPAO_GetWindow = This.OriginalIOleIPAO.GetWindow(hWnd)
	End Function
	
	Private Function IOleIPAO_ContextSensitiveHelp(ByRef This As VTableIPAODataStruct, ByVal EnterMode As Integer) As Integer
		IOleIPAO_ContextSensitiveHelp = This.OriginalIOleIPAO.ContextSensitiveHelp(EnterMode)
	End Function
	
	Private Function IOleIPAO_TranslateAccelerator(ByRef This As VTableIPAODataStruct, ByRef Msg As OLEGuids.OLEACCELMSG) As Integer
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If VarPtr(Msg) = 0 Then
			IOleIPAO_TranslateAccelerator = E_INVALIDARG
			Exit Function
		End If
		On Error GoTo CATCH_EXCEPTION
		Dim Handled As Boolean
		IOleIPAO_TranslateAccelerator = S_OK
		This.IOleIPAO.TranslateAccelerator(Handled, IOleIPAO_TranslateAccelerator, Msg.hWnd, Msg.Message, Msg.wParam, Msg.lParam, GetShiftStateFromMsg())
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If Handled = False Then IOleIPAO_TranslateAccelerator = This.OriginalIOleIPAO.TranslateAccelerator(VarPtr(Msg))
		Exit Function
CATCH_EXCEPTION: 
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		IOleIPAO_TranslateAccelerator = This.OriginalIOleIPAO.TranslateAccelerator(VarPtr(Msg))
	End Function
	
	Private Function IOleIPAO_OnFrameWindowActivate(ByRef This As VTableIPAODataStruct, ByVal Activate As Integer) As Integer
		IOleIPAO_OnFrameWindowActivate = This.OriginalIOleIPAO.OnFrameWindowActivate(Activate)
	End Function
	
	Private Function IOleIPAO_OnDocWindowActivate(ByRef This As VTableIPAODataStruct, ByVal Activate As Integer) As Integer
		IOleIPAO_OnDocWindowActivate = This.OriginalIOleIPAO.OnDocWindowActivate(Activate)
	End Function
	
	Private Function IOleIPAO_ResizeBorder(ByRef This As VTableIPAODataStruct, ByRef RC As OLEGuids.OLERECT, ByVal UIWindow As OLEGuids.IOleInPlaceUIWindow, ByVal FrameWindow As Integer) As Integer
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		IOleIPAO_ResizeBorder = This.OriginalIOleIPAO.ResizeBorder(VarPtr(RC), UIWindow, FrameWindow)
	End Function
	
	Private Function IOleIPAO_EnableModeless(ByRef This As VTableIPAODataStruct, ByVal Enable As Integer) As Integer
		IOleIPAO_EnableModeless = This.OriginalIOleIPAO.EnableModeless(Enable)
	End Function
	
	Private Sub ReplaceIOleControl(ByVal This As OLEGuids.IOleControl)
		'UPGRADE_ISSUE: ObjPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If OriginalVTableControl = 0 Then CopyMemory(OriginalVTableControl, ObjPtr(This), 4)
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		'UPGRADE_ISSUE: ObjPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		CopyMemory(ObjPtr(This), VarPtr(GetVTableControl()), 4)
	End Sub
	
	Private Sub RestoreIOleControl(ByVal This As OLEGuids.IOleControl)
		'UPGRADE_ISSUE: ObjPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If OriginalVTableControl <> 0 Then CopyMemory(ObjPtr(This), OriginalVTableControl, 4)
	End Sub
	
	Public Sub OnControlInfoChanged(ByVal This As Object, Optional ByVal OnFocus As Boolean = False)
		On Error GoTo CATCH_EXCEPTION
		Dim PropOleObject As OLEGuids.IOleObject
		Dim PropOleControlSite As OLEGuids.IOleControlSite
		PropOleObject = This
		PropOleControlSite = PropOleObject.GetClientSite
		PropOleControlSite.OnControlInfoChanged()
		If OnFocus = True Then PropOleControlSite.OnFocus(1)
CATCH_EXCEPTION: 
	End Sub
	
	Private Function GetVTableControl() As Integer
		If VTableControl(0) = 0 Then
			If OriginalVTableControl <> 0 Then
				CopyMemory(VTableControl(0), OriginalVTableControl, 12)
			Else
				'UPGRADE_WARNING: Add a delegate for AddressOf IOleControl_QueryInterface Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
				VTableControl(0) = ProcPtr(AddressOf IOleControl_QueryInterface)
				'UPGRADE_WARNING: Add a delegate for AddressOf IOleControl_AddRef Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
				VTableControl(1) = ProcPtr(AddressOf IOleControl_AddRef)
				'UPGRADE_WARNING: Add a delegate for AddressOf IOleControl_Release Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
				VTableControl(2) = ProcPtr(AddressOf IOleControl_Release)
			End If
			'UPGRADE_WARNING: Add a delegate for AddressOf IOleControl_GetControlInfo Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableControl(3) = ProcPtr(AddressOf IOleControl_GetControlInfo)
			'UPGRADE_WARNING: Add a delegate for AddressOf IOleControl_OnMnemonic Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableControl(4) = ProcPtr(AddressOf IOleControl_OnMnemonic)
			'UPGRADE_WARNING: Add a delegate for AddressOf IOleControl_OnAmbientPropertyChange Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableControl(5) = ProcPtr(AddressOf IOleControl_OnAmbientPropertyChange)
			If OriginalVTableControl <> 0 Then
				CopyMemory(VTableControl(6), UnsignedAdd(OriginalVTableControl, 24), 4)
			Else
				'UPGRADE_WARNING: Add a delegate for AddressOf IOleControl_FreezeEvents Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
				VTableControl(6) = ProcPtr(AddressOf IOleControl_FreezeEvents)
			End If
		End If
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		GetVTableControl = VarPtr(VTableControl(0))
	End Function
	
	Private Function IOleControl_QueryInterface(ByRef This As Integer, ByRef IID As OLEGuids.OLECLSID, ByRef pvObj As Integer) As Integer
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If VarPtr(pvObj) = 0 Then
			IOleControl_QueryInterface = E_POINTER
			Exit Function
		End If
		Dim IUnk As OLEGuids.IUnknownUnrestricted
		If OriginalVTableControl <> 0 Then
			This = OriginalVTableControl
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object IUnk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(IUnk, VarPtr(This), 4)
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			IOleControl_QueryInterface = IUnk.QueryInterface(VarPtr(IID), pvObj)
			'UPGRADE_WARNING: Couldn't resolve default property of object IUnk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(IUnk, 0, 4)
			This = GetVTableControl()
		End If
	End Function
	
	Private Function IOleControl_AddRef(ByRef This As Integer) As Integer
		Dim IUnk As OLEGuids.IUnknownUnrestricted
		If OriginalVTableControl <> 0 Then
			This = OriginalVTableControl
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object IUnk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(IUnk, VarPtr(This), 4)
			IOleControl_AddRef = IUnk.AddRef()
			'UPGRADE_WARNING: Couldn't resolve default property of object IUnk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(IUnk, 0, 4)
			This = GetVTableControl()
		End If
	End Function
	
	Private Function IOleControl_Release(ByRef This As Integer) As Integer
		Dim IUnk As OLEGuids.IUnknownUnrestricted
		If OriginalVTableControl <> 0 Then
			This = OriginalVTableControl
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object IUnk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(IUnk, VarPtr(This), 4)
			IOleControl_Release = IUnk.Release()
			'UPGRADE_WARNING: Couldn't resolve default property of object IUnk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(IUnk, 0, 4)
			This = GetVTableControl()
		End If
	End Function
	
	Private Function IOleControl_GetControlInfo(ByRef This As Integer, ByRef CI As OLEGuids.OLECONTROLINFO) As Integer
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If VarPtr(CI) = 0 Then
			IOleControl_GetControlInfo = E_POINTER
			Exit Function
		End If
		On Error GoTo CATCH_EXCEPTION
		Dim ShadowIOleControlVB As OLEGuids.IOleControlVB
		Dim Handled As Boolean
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		ShadowIOleControlVB = PtrToObj(VarPtr(This))
		'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		CI.cb = LenB(CI)
		ShadowIOleControlVB.GetControlInfo(Handled, (CI.cAccel), (CI.hAccel), (CI.dwFlags))
		If Handled = False Then
			IOleControl_GetControlInfo = Original_IOleControl_GetControlInfo(This, CI)
		Else
			If CI.cAccel > 0 And CI.hAccel = 0 Then
				IOleControl_GetControlInfo = E_OUTOFMEMORY
			Else
				IOleControl_GetControlInfo = S_OK
			End If
		End If
		Exit Function
CATCH_EXCEPTION: 
		IOleControl_GetControlInfo = Original_IOleControl_GetControlInfo(This, CI)
	End Function
	
	Private Function IOleControl_OnMnemonic(ByRef This As Integer, ByRef Msg As OLEGuids.OLEACCELMSG) As Integer
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If VarPtr(Msg) = 0 Then
			IOleControl_OnMnemonic = E_INVALIDARG
			Exit Function
		End If
		On Error GoTo CATCH_EXCEPTION
		Dim ShadowIOleControlVB As OLEGuids.IOleControlVB
		Dim Handled As Boolean
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		ShadowIOleControlVB = PtrToObj(VarPtr(This))
		ShadowIOleControlVB.OnMnemonic(Handled, Msg.hWnd, Msg.Message, Msg.wParam, Msg.lParam, GetShiftStateFromMsg())
		If Handled = False Then
			IOleControl_OnMnemonic = Original_IOleControl_OnMnemonic(This, Msg)
		Else
			IOleControl_OnMnemonic = S_OK
		End If
		Exit Function
CATCH_EXCEPTION: 
		IOleControl_OnMnemonic = Original_IOleControl_OnMnemonic(This, Msg)
	End Function
	
	Private Function IOleControl_OnAmbientPropertyChange(ByRef This As Integer, ByVal DispID As Integer) As Integer
		IOleControl_OnAmbientPropertyChange = Original_IOleControl_OnAmbientPropertyChange(This, DispID)
	End Function
	
	Private Function IOleControl_FreezeEvents(ByRef This As Integer, ByVal bFreeze As Integer) As Integer
		IOleControl_FreezeEvents = Original_IOleControl_FreezeEvents(This, bFreeze)
	End Function
	
	Private Function Original_IOleControl_GetControlInfo(ByRef This As Integer, ByRef CI As OLEGuids.OLECONTROLINFO) As Integer
		Dim ShadowIOleControl As OLEGuids.IOleControl
		If OriginalVTableControl <> 0 Then
			This = OriginalVTableControl
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIOleControl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIOleControl, VarPtr(This), 4)
			Original_IOleControl_GetControlInfo = ShadowIOleControl.GetControlInfo(CI)
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIOleControl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIOleControl, 0, 4)
			This = GetVTableControl()
		Else
			Original_IOleControl_GetControlInfo = E_NOTIMPL
		End If
	End Function
	
	Private Function Original_IOleControl_OnMnemonic(ByRef This As Integer, ByRef Msg As OLEGuids.OLEACCELMSG) As Integer
		Dim ShadowIOleControl As OLEGuids.IOleControl
		If OriginalVTableControl <> 0 Then
			This = OriginalVTableControl
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIOleControl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIOleControl, VarPtr(This), 4)
			Original_IOleControl_OnMnemonic = ShadowIOleControl.OnMnemonic(Msg)
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIOleControl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIOleControl, 0, 4)
			This = GetVTableControl()
		Else
			Original_IOleControl_OnMnemonic = E_NOTIMPL
		End If
	End Function
	
	Private Function Original_IOleControl_OnAmbientPropertyChange(ByRef This As Integer, ByVal DispID As Integer) As Integer
		Dim ShadowIOleControl As OLEGuids.IOleControl
		If OriginalVTableControl <> 0 Then
			This = OriginalVTableControl
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIOleControl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIOleControl, VarPtr(This), 4)
			ShadowIOleControl.OnAmbientPropertyChange(DispID)
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIOleControl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIOleControl, 0, 4)
			This = GetVTableControl()
		End If
		' This function returns S_OK in all cases.
		Original_IOleControl_OnAmbientPropertyChange = S_OK
	End Function
	
	Private Function Original_IOleControl_FreezeEvents(ByRef This As Integer, ByVal bFreeze As Integer) As Integer
		Dim ShadowIOleControl As OLEGuids.IOleControl
		If OriginalVTableControl <> 0 Then
			This = OriginalVTableControl
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIOleControl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIOleControl, VarPtr(This), 4)
			ShadowIOleControl.FreezeEvents(bFreeze)
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIOleControl. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIOleControl, 0, 4)
			This = GetVTableControl()
		End If
		' This function returns S_OK in all cases.
		Original_IOleControl_FreezeEvents = S_OK
	End Function
	
	Private Sub ReplaceIPPB(ByVal This As OLEGuids.IPerPropertyBrowsing)
		'UPGRADE_ISSUE: ObjPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If OriginalVTablePPB = 0 Then CopyMemory(OriginalVTablePPB, ObjPtr(This), 4)
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		'UPGRADE_ISSUE: ObjPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		CopyMemory(ObjPtr(This), VarPtr(GetVTablePPB()), 4)
	End Sub
	
	Private Sub RestoreIPPB(ByVal This As OLEGuids.IPerPropertyBrowsing)
		'UPGRADE_ISSUE: ObjPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If OriginalVTablePPB <> 0 Then CopyMemory(ObjPtr(This), OriginalVTablePPB, 4)
	End Sub
	
	Public Function GetDispID(ByVal This As Object, ByRef MethodName As String) As Integer
		'UPGRADE_WARNING: Arrays in structure IID_NULL may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim IDispatch As OLEGuids.IDispatch
		Dim IID_NULL As OLEGuids.OLECLSID
		IDispatch = This
		IDispatch.GetIDsOfNames(IID_NULL, MethodName, 1, 0, GetDispID)
	End Function
	
	Private Function GetVTablePPB() As Integer
		If VTablePPB(0) = 0 Then
			If OriginalVTablePPB <> 0 Then
				CopyMemory(VTablePPB(0), OriginalVTablePPB, 12)
			Else
				'UPGRADE_WARNING: Add a delegate for AddressOf IPPB_QueryInterface Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
				VTablePPB(0) = ProcPtr(AddressOf IPPB_QueryInterface)
				'UPGRADE_WARNING: Add a delegate for AddressOf IPPB_AddRef Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
				VTablePPB(1) = ProcPtr(AddressOf IPPB_AddRef)
				'UPGRADE_WARNING: Add a delegate for AddressOf IPPB_Release Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
				VTablePPB(2) = ProcPtr(AddressOf IPPB_Release)
			End If
			'UPGRADE_WARNING: Add a delegate for AddressOf IPPB_GetDisplayString Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTablePPB(3) = ProcPtr(AddressOf IPPB_GetDisplayString)
			If OriginalVTablePPB <> 0 Then
				CopyMemory(VTablePPB(4), UnsignedAdd(OriginalVTablePPB, 16), 4)
			Else
				'UPGRADE_WARNING: Add a delegate for AddressOf IPPB_MapPropertyToPage Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
				VTablePPB(4) = ProcPtr(AddressOf IPPB_MapPropertyToPage)
			End If
			'UPGRADE_WARNING: Add a delegate for AddressOf IPPB_GetPredefinedStrings Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTablePPB(5) = ProcPtr(AddressOf IPPB_GetPredefinedStrings)
			'UPGRADE_WARNING: Add a delegate for AddressOf IPPB_GetPredefinedValue Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTablePPB(6) = ProcPtr(AddressOf IPPB_GetPredefinedValue)
		End If
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		GetVTablePPB = VarPtr(VTablePPB(0))
	End Function
	
	Private Function IPPB_QueryInterface(ByRef This As Integer, ByRef IID As OLEGuids.OLECLSID, ByRef pvObj As Integer) As Integer
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If VarPtr(pvObj) = 0 Then
			IPPB_QueryInterface = E_POINTER
			Exit Function
		End If
		Dim IUnk As OLEGuids.IUnknownUnrestricted
		If OriginalVTablePPB <> 0 Then
			This = OriginalVTablePPB
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object IUnk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(IUnk, VarPtr(This), 4)
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			IPPB_QueryInterface = IUnk.QueryInterface(VarPtr(IID), pvObj)
			'UPGRADE_WARNING: Couldn't resolve default property of object IUnk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(IUnk, 0, 4)
			This = GetVTablePPB()
		End If
	End Function
	
	Private Function IPPB_AddRef(ByRef This As Integer) As Integer
		Dim IUnk As OLEGuids.IUnknownUnrestricted
		If OriginalVTablePPB <> 0 Then
			This = OriginalVTablePPB
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object IUnk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(IUnk, VarPtr(This), 4)
			IPPB_AddRef = IUnk.AddRef()
			'UPGRADE_WARNING: Couldn't resolve default property of object IUnk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(IUnk, 0, 4)
			This = GetVTablePPB()
		End If
	End Function
	
	Private Function IPPB_Release(ByRef This As Integer) As Integer
		Dim IUnk As OLEGuids.IUnknownUnrestricted
		If OriginalVTablePPB <> 0 Then
			This = OriginalVTablePPB
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object IUnk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(IUnk, VarPtr(This), 4)
			IPPB_Release = IUnk.Release()
			'UPGRADE_WARNING: Couldn't resolve default property of object IUnk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(IUnk, 0, 4)
			This = GetVTablePPB()
		End If
	End Function
	
	Private Function IPPB_GetDisplayString(ByRef This As Integer, ByVal DispID As Integer, ByRef pDisplayName As Integer) As Integer
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If VarPtr(pDisplayName) = 0 Then
			IPPB_GetDisplayString = E_POINTER
			Exit Function
		End If
		On Error GoTo CATCH_EXCEPTION
		Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB
		Dim Handled As Boolean
		Dim DisplayName As String
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		ShadowIPerPropertyBrowsingVB = PtrToObj(VarPtr(This))
		ShadowIPerPropertyBrowsingVB.GetDisplayString(Handled, DispID, DisplayName)
		If Handled = False Then
			IPPB_GetDisplayString = Original_IPPB_GetDisplayString(This, DispID, pDisplayName)
		Else
			'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			pDisplayName = SysAllocString(StrPtr(DisplayName))
			IPPB_GetDisplayString = S_OK
		End If
		Exit Function
CATCH_EXCEPTION: 
		IPPB_GetDisplayString = Original_IPPB_GetDisplayString(This, DispID, pDisplayName)
	End Function
	
	Private Function IPPB_MapPropertyToPage(ByRef This As Integer, ByVal DispID As Integer, ByRef pCLSID As OLEGuids.OLECLSID) As Integer
		IPPB_MapPropertyToPage = Original_IPPB_MapPropertyToPage(This, DispID, pCLSID)
	End Function
	
	Private Function IPPB_GetPredefinedStrings(ByRef This As Integer, ByVal DispID As Integer, ByRef pCaStringsOut As OLEGuids.OLECALPOLESTR, ByRef pCaCookiesOut As OLEGuids.OLECADWORD) As Integer
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If VarPtr(pCaStringsOut) = 0 Or VarPtr(pCaCookiesOut) = 0 Then
			IPPB_GetPredefinedStrings = E_POINTER
			Exit Function
		End If
		On Error GoTo CATCH_EXCEPTION
		Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB
		Dim Handled As Boolean
		ReDim StringsOutArray(0)
		ReDim CookiesOutArray(0)
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		ShadowIPerPropertyBrowsingVB = PtrToObj(VarPtr(This))
		ShadowIPerPropertyBrowsingVB.GetPredefinedStrings(Handled, DispID, StringsOutArray, CookiesOutArray)
		Dim pElems, cElems, nElemCount As Integer
		Dim lpString As Integer
		If Handled = False Or UBound(StringsOutArray) = 0 Then
			IPPB_GetPredefinedStrings = Original_IPPB_GetPredefinedStrings(This, DispID, pCaStringsOut, pCaCookiesOut)
		Else
			cElems = UBound(StringsOutArray)
			If Not UBound(CookiesOutArray) = cElems Then ReDim Preserve CookiesOutArray(cElems)
			pElems = CoTaskMemAlloc(cElems * 4)
			pCaStringsOut.cElems = cElems
			pCaStringsOut.pElems = pElems
			For nElemCount = 0 To cElems - 1
				lpString = CoTaskMemAlloc(Len(StringsOutArray(nElemCount)) + 1)
				'UPGRADE_ISSUE: StrPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				CopyMemory(lpString, StrPtr(StringsOutArray(nElemCount)), 4)
				CopyMemory(UnsignedAdd(pElems, nElemCount * 4), lpString, 4)
			Next nElemCount
			pElems = CoTaskMemAlloc(cElems * 4)
			pCaCookiesOut.cElems = cElems
			pCaCookiesOut.pElems = pElems
			For nElemCount = 0 To cElems - 1
				CopyMemory(UnsignedAdd(pElems, nElemCount * 4), CookiesOutArray(nElemCount), 4)
			Next nElemCount
			IPPB_GetPredefinedStrings = S_OK
		End If
		Exit Function
CATCH_EXCEPTION: 
		IPPB_GetPredefinedStrings = Original_IPPB_GetPredefinedStrings(This, DispID, pCaStringsOut, pCaCookiesOut)
	End Function
	
	Private Function IPPB_GetPredefinedValue(ByRef This As Integer, ByVal DispID As Integer, ByVal dwCookie As Integer, ByRef pVarOut As Object) As Integer
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If VarPtr(pVarOut) = 0 Then
			IPPB_GetPredefinedValue = E_POINTER
			Exit Function
		End If
		On Error GoTo CATCH_EXCEPTION
		Dim ShadowIPerPropertyBrowsingVB As OLEGuids.IPerPropertyBrowsingVB
		Dim Handled As Boolean
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		ShadowIPerPropertyBrowsingVB = PtrToObj(VarPtr(This))
		ShadowIPerPropertyBrowsingVB.GetPredefinedValue(Handled, DispID, dwCookie, pVarOut)
		If Handled = False Then
			IPPB_GetPredefinedValue = Original_IPPB_GetPredefinedValue(This, DispID, dwCookie, pVarOut)
		Else
			IPPB_GetPredefinedValue = S_OK
		End If
		Exit Function
CATCH_EXCEPTION: 
		IPPB_GetPredefinedValue = Original_IPPB_GetPredefinedValue(This, DispID, dwCookie, pVarOut)
	End Function
	
	Private Function Original_IPPB_GetDisplayString(ByRef This As Integer, ByVal DispID As Integer, ByRef pDisplayName As Integer) As Integer
		Dim ShadowIPPB As OLEGuids.IPerPropertyBrowsing
		If OriginalVTablePPB <> 0 Then
			This = OriginalVTablePPB
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIPPB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIPPB, VarPtr(This), 4)
			Original_IPPB_GetDisplayString = ShadowIPPB.GetDisplayString(DispID, pDisplayName)
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIPPB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIPPB, 0, 4)
			This = GetVTablePPB()
		End If
	End Function
	
	Private Function Original_IPPB_MapPropertyToPage(ByRef This As Integer, ByVal DispID As Integer, ByRef pCLSID As OLEGuids.OLECLSID) As Integer
		Dim ShadowIPPB As OLEGuids.IPerPropertyBrowsing
		If OriginalVTablePPB <> 0 Then
			This = OriginalVTablePPB
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIPPB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIPPB, VarPtr(This), 4)
			Original_IPPB_MapPropertyToPage = ShadowIPPB.MapPropertyToPage(DispID, pCLSID)
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIPPB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIPPB, 0, 4)
			This = GetVTablePPB()
		End If
	End Function
	
	Private Function Original_IPPB_GetPredefinedStrings(ByRef This As Integer, ByVal DispID As Integer, ByRef pCaStringsOut As OLEGuids.OLECALPOLESTR, ByRef pCaCookiesOut As OLEGuids.OLECADWORD) As Integer
		Dim ShadowIPPB As OLEGuids.IPerPropertyBrowsing
		If OriginalVTablePPB <> 0 Then
			This = OriginalVTablePPB
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIPPB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIPPB, VarPtr(This), 4)
			Original_IPPB_GetPredefinedStrings = ShadowIPPB.GetPredefinedStrings(DispID, pCaStringsOut, pCaCookiesOut)
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIPPB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIPPB, 0, 4)
			This = GetVTablePPB()
		End If
	End Function
	
	Private Function Original_IPPB_GetPredefinedValue(ByRef This As Integer, ByVal DispID As Integer, ByVal dwCookie As Integer, ByRef pVarOut As Object) As Integer
		Dim ShadowIPPB As OLEGuids.IPerPropertyBrowsing
		If OriginalVTablePPB <> 0 Then
			This = OriginalVTablePPB
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIPPB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIPPB, VarPtr(This), 4)
			Original_IPPB_GetPredefinedValue = ShadowIPPB.GetPredefinedValue(DispID, dwCookie, pVarOut)
			'UPGRADE_WARNING: Couldn't resolve default property of object ShadowIPPB. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(ShadowIPPB, 0, 4)
			This = GetVTablePPB()
		End If
	End Function
	
	Public Function GetNewEnum(ByVal This As Object, ByVal Upper As Integer, ByVal Lower As Integer) As stdole.IEnumVARIANT
		Dim VTableEnumVARIANTData As VTableEnumVARIANTDataStruct
		Dim hMem As Integer
		With VTableEnumVARIANTData
			.VTable = GetVTableEnumVARIANT()
			.RefCount = 1
			.Enumerable = This
			.Index = Lower
			.Count = Upper
			'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			hMem = CoTaskMemAlloc(LenB(VTableEnumVARIANTData))
			If hMem <> 0 Then
				'UPGRADE_ISSUE: LenB function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object VTableEnumVARIANTData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CopyMemory(hMem, VTableEnumVARIANTData, LenB(VTableEnumVARIANTData))
				'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				CopyMemory(VarPtr(GetNewEnum), hMem, 4)
				'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				CopyMemory(VarPtr(.Enumerable), 0, 4)
			End If
		End With
	End Function
	
	Private Function GetVTableEnumVARIANT() As Integer
		If VTableEnumVARIANT(0) = 0 Then
			'UPGRADE_WARNING: Add a delegate for AddressOf IEnumVARIANT_QueryInterface Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableEnumVARIANT(0) = ProcPtr(AddressOf IEnumVARIANT_QueryInterface)
			'UPGRADE_WARNING: Add a delegate for AddressOf IEnumVARIANT_AddRef Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableEnumVARIANT(1) = ProcPtr(AddressOf IEnumVARIANT_AddRef)
			'UPGRADE_WARNING: Add a delegate for AddressOf IEnumVARIANT_Release Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableEnumVARIANT(2) = ProcPtr(AddressOf IEnumVARIANT_Release)
			'UPGRADE_WARNING: Add a delegate for AddressOf IEnumVARIANT_Next Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableEnumVARIANT(3) = ProcPtr(AddressOf IEnumVARIANT_Next)
			'UPGRADE_WARNING: Add a delegate for AddressOf IEnumVARIANT_Skip Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableEnumVARIANT(4) = ProcPtr(AddressOf IEnumVARIANT_Skip)
			'UPGRADE_WARNING: Add a delegate for AddressOf IEnumVARIANT_Reset Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableEnumVARIANT(5) = ProcPtr(AddressOf IEnumVARIANT_Reset)
			'UPGRADE_WARNING: Add a delegate for AddressOf IEnumVARIANT_Clone Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="E9E157F7-EF0C-4016-87B7-7D7FBBC6EE08"'
			VTableEnumVARIANT(6) = ProcPtr(AddressOf IEnumVARIANT_Clone)
		End If
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		GetVTableEnumVARIANT = VarPtr(VTableEnumVARIANT(0))
	End Function
	
	Private Function IEnumVARIANT_QueryInterface(ByRef This As VTableEnumVARIANTDataStruct, ByRef IID As OLEGuids.OLECLSID, ByRef pvObj As Integer) As Integer
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If VarPtr(pvObj) = 0 Then
			IEnumVARIANT_QueryInterface = E_POINTER
			Exit Function
		End If
		' IID_IEnumVARIANT = {00020404-0000-0000-C000-000000000046}
		If IID.Data1 = &H20404 And IID.Data2 = &H0 And IID.Data3 = &H0 Then
			If IID.Data4(0) = &HC0 And IID.Data4(1) = &H0 And IID.Data4(2) = &H0 And IID.Data4(3) = &H0 And IID.Data4(4) = &H0 And IID.Data4(5) = &H0 And IID.Data4(6) = &H0 And IID.Data4(7) = &H46 Then
				'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
				pvObj = VarPtr(This)
				IEnumVARIANT_AddRef(This)
				IEnumVARIANT_QueryInterface = S_OK
			Else
				IEnumVARIANT_QueryInterface = E_NOINTERFACE
			End If
		Else
			IEnumVARIANT_QueryInterface = E_NOINTERFACE
		End If
	End Function
	
	Private Function IEnumVARIANT_AddRef(ByRef This As VTableEnumVARIANTDataStruct) As Integer
		This.RefCount = This.RefCount + 1
		IEnumVARIANT_AddRef = This.RefCount
	End Function
	
	Private Function IEnumVARIANT_Release(ByRef This As VTableEnumVARIANTDataStruct) As Integer
		This.RefCount = This.RefCount - 1
		IEnumVARIANT_Release = This.RefCount
		If IEnumVARIANT_Release = 0 Then
			'UPGRADE_NOTE: Object This.Enumerable may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			This.Enumerable = Nothing
			'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
			CoTaskMemFree(VarPtr(This))
		End If
	End Function
	
	Private Function IEnumVARIANT_Next(ByRef This As VTableEnumVARIANTDataStruct, ByVal VntCount As Integer, ByVal VntArrPtr As Integer, ByRef pcvFetched As Integer) As Integer
		If VntArrPtr = 0 Then
			IEnumVARIANT_Next = E_INVALIDARG
			Exit Function
		End If
		On Error GoTo CATCH_EXCEPTION
		Const VARIANT_CB As Integer = 16
		Dim Fetched As Integer
		With This
			Do Until .Index > .Count
				VariantCopyToPtr(VntArrPtr, .Enumerable(.Index))
				.Index = .Index + 1
				Fetched = Fetched + 1
				If Fetched = VntCount Then Exit Do
				VntArrPtr = UnsignedAdd(VntArrPtr, VARIANT_CB)
			Loop 
		End With
		If Fetched = VntCount Then
			IEnumVARIANT_Next = S_OK
		Else
			IEnumVARIANT_Next = S_FALSE
		End If
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If VarPtr(pcvFetched) <> 0 Then pcvFetched = Fetched
		Exit Function
CATCH_EXCEPTION: 
		'UPGRADE_ISSUE: VarPtr function is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="367764E5-F3F8-4E43-AC3E-7FE0B5E074E2"'
		If VarPtr(pcvFetched) <> 0 Then pcvFetched = 0
		IEnumVARIANT_Next = E_NOTIMPL
	End Function
	
	Private Function IEnumVARIANT_Skip(ByRef This As VTableEnumVARIANTDataStruct, ByVal VntCount As Integer) As Integer
		IEnumVARIANT_Skip = E_NOTIMPL
	End Function
	
	Private Function IEnumVARIANT_Reset(ByRef This As VTableEnumVARIANTDataStruct) As Integer
		IEnumVARIANT_Reset = E_NOTIMPL
	End Function
	
	Private Function IEnumVARIANT_Clone(ByRef This As VTableEnumVARIANTDataStruct, ByRef ppEnum As stdole.IEnumVARIANT) As Integer
		IEnumVARIANT_Clone = E_NOTIMPL
	End Function
End Module