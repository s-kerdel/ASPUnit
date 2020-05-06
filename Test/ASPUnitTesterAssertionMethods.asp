<% Option Explicit %>

<!-- #include file="../Lib/ASPUnit.asp" -->

<%
	Dim objLifecycle
	Set objLifecycle = ASPUnit.CreateLifeCycle("Setup", "Teardown")

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester Ok Assertion Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterOkPassedTruthy"), _
				ASPUnit.CreateTest("ASPUnitTesterOkPassedFalsey") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester Equal Assertion Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterEqualPassedTruthy"), _
				ASPUnit.CreateTest("ASPUnitTesterEqualPassedFalsey") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester NotEqual Assertion Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterNotEqualPassedTruthy"), _
				ASPUnit.CreateTest("ASPUnitTesterNotEqualPassedFalsey") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester Same Assertion Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterSamePassedTruthy"), _
				ASPUnit.CreateTest("ASPUnitTesterSamePassedFalsey") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester NotSame Assertion Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterNotSamePassedTruthy"), _
				ASPUnit.CreateTest("ASPUnitTesterNotSamePassedFalsey") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.AddModule( _
		ASPUnit.CreateModule( _
			"ASPUnitTester EqualDictionaries Assertion Method Tests", _
			Array( _
				ASPUnit.CreateTest("ASPUnitTesterEqualDictionariesPassedTruthy"), _
				ASPUnit.CreateTest("ASPUnitTesterEqualDictionariesPassedFalsy") _
			), _
			objLifecycle _
		) _
	)

	Call ASPUnit.Run()

	' Create a global instance of ASPUnitTester for testing

	Sub Setup()
		Call ExecuteGlobal("Dim objService")
		Set objService = New ASPUnitTester
	End Sub

	Sub Teardown()
		Set objService = Nothing
	End Sub

	' Ok Assertion Method Tests

	Sub ASPUnitTesterOkPassedTruthy()
		Call ASPUnit.Equal(objService.Ok(True, ""), True, "Ok method should return truthy")
	End Sub

	Sub ASPUnitTesterOkPassedFalsey()
		Call ASPUnit.Equal(objService.Ok(False, ""), False, "Ok method should return falsey")
	End Sub

	' Equal Assertion Method Tests

	Sub ASPUnitTesterEqualPassedTruthy()
		Call ASPUnit.Equal(objService.Equal(True, True, ""), True, "Equal method should return truthy with equal values")
	End Sub

	Sub ASPUnitTesterEqualPassedFalsey()
		Call ASPUnit.Equal(objService.Equal(True, False, ""), False, "Equal method should return falsey with unequal values")
	End Sub

	' NotEqual Assertion Method Tests

	Sub ASPUnitTesterNotEqualPassedTruthy()
		Call ASPUnit.Equal(objService.NotEqual(True, False, ""), True, "NotEqual method should return truthy with unequal values")
	End Sub

	Sub ASPUnitTesterNotEqualPassedFalsey()
		Call ASPUnit.Equal(objService.NotEqual(True, True, ""), False, "NotEqual method should return falsey with equal values")
	End Sub

	' Same Assertion Method Tests

	Sub ASPUnitTesterSamePassedTruthy()
		Dim objA, _
			objB

		Set objA = New RegExp
		Set objB = objA

		Call ASPUnit.Equal(objService.Same(objA, objB, ""), True, "Same method should return truthy with same references")

		Set objB = Nothing
		Set objA = Nothing
	End Sub

	Sub ASPUnitTesterSamePassedFalsey()
		Dim objA, _
			objB

		Set objA = New RegExp
		Set objB = New RegExp

		Call ASPUnit.Equal(objService.Same(objA, objB, ""), False, "Same method should return falsey with different references")

		Set objB = Nothing
		Set objA = Nothing
	End Sub

	' NotSame Assertion Method Tests

	Sub ASPUnitTesterNotSamePassedTruthy()
		Dim objA, _
			objB

		Set objA = New RegExp
		Set objB = New RegExp

		Call ASPUnit.Equal(objService.NotSame(objA, objB, ""), True, "NotSame method should return truthy with different references")

		Set objB = Nothing
		Set objA = Nothing
	End Sub

	Sub ASPUnitTesterNotSamePassedFalsey()
		Dim objA, _
			objB

		Set objA = New RegExp
		Set objB = objA

		Call ASPUnit.Equal(objService.NotSame(objA, objB, ""), False, "NotSame method should return falsey with same references")

		Set objB = Nothing
		Set objA = Nothing
	End Sub

	' EqualDictionaries Assertion Method Tests
	Sub ASPUnitTesterEqualDictionariesPassedTruthy()
		Dim dic1 : set dic1 = Server.CreateObject("Scripting.Dictionary")
		Dim dic2 : set dic2 = Server.CreateObject("Scripting.Dictionary")
		dic1.add "key1", "value1"
		dic2.add "key1", "value1"
		Call ASPUnit.Equal(objService.EqualDictionaries(dic1, dic2, ""), True, "EqualDictionaries method should return truthy with equal dictionaries")
		
		Set dic1 = Nothing
		Set dic2 = Nothing
	End Sub
	Sub ASPUnitTesterEqualDictionariesPassedFalsy()
		Dim dic1 : set dic1 = Server.CreateObject("Scripting.Dictionary")
		Dim dic2 : set dic2 = Server.CreateObject("Scripting.Dictionary")
		dic1.add "key1", "value1"
		dic2.add "key2", "value2"
		Call ASPUnit.Equal(objService.EqualDictionaries(dic1, dic2, ""), False, "EqualDictionaries method should return falsy with not equal dictionaries")
		
		Set dic1 = Nothing
		Set dic2 = Nothing
	End Sub
%>