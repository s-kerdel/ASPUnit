<%
	Class ASPUnitTester
		Private _
			m_Responder, _
			m_Scenario

		Private _
			m_CurrentModule, _
			m_CurrentTest

		Private Sub Class_Initialize()
			Set m_Responder = New ASPUnitJSONResponder
			Set m_Scenario = New ASPUnitScenario
		End Sub

		Private Sub Class_Terminate()
			Set m_Scenario = Nothing
			Set m_Responder = Nothing
		End Sub

		Public Property Set Responder(ByRef objValue)
			Set m_Responder = objValue
		End Property

		Public Property Get Modules()
			Set Modules = m_Scenario.Modules
		End Property

		' Factory methods for private classes

		Public Function CreateModule(strName, arrTests, objLifecycle)
			Dim objReturn, _
				i

			Set objReturn = New ASPUnitModule
			objReturn.Name = strName
			For i = 0 To UBound(arrTests)
				objReturn.Tests.Add(arrTests(i))
			Next
			Set objReturn.Lifecycle = objLifecycle

			Set CreateModule = objReturn
		End Function

		Public Function CreateTest(strName)
			Dim objReturn

			Set objReturn = New ASPUnitTest
			objReturn.Name = strName

			Set CreateTest = objReturn
		End Function

		Public Function CreateLifecycle(strSetup, strTeardown)
			Dim objReturn

			Set objReturn = New ASPUnitTestLifecycle
			objReturn.Setup = strSetup
			objReturn.Teardown = strTeardown

			Set CreateLifecycle = objReturn
		End Function

		' Public methods to add modules

		Public Sub AddModule(objModule)
			Call m_Scenario.Modules.Add(objModule)
		End Sub

		Public Sub AddModules(arrModules)
			Dim i

			For i = 0 To UBound(arrModules)
				Call AddModule(arrModules(i))
			Next
		End Sub

		' Assertion Methods

		Private Function Assert(blnResult, varActual, varExpected, strDescription)
			If IsObject(m_CurrentTest) Then
				' Register every Assert call to the current test.
				Dim iAssertIndex : iAssertIndex = m_CurrentTest.Assertions.Count
				m_CurrentTest.Assertions.Add iAssertIndex, New ASPUnitTestAssertion
				m_CurrentTest.Assertions(iAssertIndex).Passed = blnResult
				m_CurrentTest.Assertions(iAssertIndex).Description = strDescription

				If IsObject(varActual) Then ' Objects require `Set`.
					Set m_CurrentTest.Assertions(iAssertIndex).Actual = varActual
				Else	m_CurrentTest.Assertions(iAssertIndex).Actual = varActual
				End If
				
				If IsObject(varExpected) Then ' Objects require `Set`.
					Set m_CurrentTest.Assertions(iAssertIndex).Expected = varExpected
				Else    m_CurrentTest.Assertions(iAssertIndex).Expected = varExpected
				End If

				m_CurrentTest.Assertions(iAssertIndex).Actual = varActual
				m_CurrentTest.Assertions(iAssertIndex).Expected = varExpected
				' Passed (Unit) will mark result red if one or more assertions failed.
				m_CurrentTest.Passed = m_CurrentTest.Passed and blnResult
			End If

			Assert = blnResult
		End Function

		Public Function Ok(blnResult, strDescription)
			Ok = blnResult
			Assert blnResult, blnResult, True, strDescription
		End Function

		Public Function Equal(varActual, varExpected, strDescription)
			' Null variables fail using the '=' operator.
			If isNull(varActual) OR isNull(varExpected) Then
				Equal = (IsNull(varActual) = isNull(varExpected))
			Else
				Equal = (varActual = varExpected)
			End If
			Assert Equal, varActual, varExpected, strDescription
		End Function

		Public Function StrictEqual(varActual, varExpected, strDescription)
			If TypeName(varActual) = TypeName(varExpected) Then
				' Null variables fail using the '=' operator.
				If isNull(varActual) OR isNull(varExpected) Then
					StrictEqual = (IsNull(varActual) = isNull(varExpected))
				Else
					StrictEqual = (varActual = varExpected)
				End If
			Else
				StrictEqual = False
			End If
			Assert StrictEqual, varActual, varExpected, strDescription
		End Function

		Public Function NotEqual(varActual, varExpected, strDescription)
			' Null variables fail using the '=' operator.
			If isNull(varActual) OR isNull(varExpected) Then
				NotEqual = Not (IsNull(varActual) = isNull(varExpected))
			Else
				NotEqual = Not (varActual = varExpected)
			End If
			Assert NotEqual, varActual, varExpected, strDescription
		End Function

		Public Function Same(varActual, varExpected, strDescription)
			Same = (varActual Is varExpected)
			Assert Same, varActual, varExpected, strDescription
		End Function

		Public Function NotSame(varActual, varExpected, strDescription)
			NotSame = Not (varActual Is varExpected)
			Assert NotSame, varActual, varExpected, strDescription
		End Function

		Public Function InstanceOf(objToCheck, strExpectedType, strDescription)
			InstanceOf = Equal(TypeName(objToCheck), strExpectedType, strDescription)
		End Function

		Public Function EqualDictionaries(varActual, varExpected, strDescription)
			EqualDictionaries = matchDictionaries(varActual, varExpected)
			Assert EqualDictionaries, varActual, varExpected, strDescription
		End Function

		' Methods to run module tests

		Public Sub Run()
			Dim objModule, _
				i

			For i = 0 To (m_Scenario.Modules.Count - 1)
				Set objModule = m_Scenario.Modules.Item(i)
				Call RunModule(objModule)

				m_Scenario.TestCount = m_Scenario.TestCount + objModule.TestCount
				m_Scenario.PassCount = m_Scenario.PassCount + objModule.PassCount
				m_Scenario.FailCount = m_Scenario.FailCount + objModule.FailCount

				Set objModule = Nothing
			Next

			m_Responder.Respond(m_Scenario)
		End Sub

		Private Sub RunModule(ByRef objModule)
			Dim intTimeStart, _
				intTimeEnd, _
				objTest, _
				i

			Set m_CurrentModule = objModule

			intTimeStart = Timer
			For i = 0 To (objModule.Tests.Count - 1)
				Set objTest = objModule.Tests.Item(i)

				Call RunTestModuleSetup(objModule)
				Call RunModuleTest(objTest)
				Call RunTestModuleTeardown(objModule)

				Dim iAssertIndex
				For iAssertIndex = 0 To (objTest.Assertions.Count - 1)
					objModule.TestCount = objModule.TestCount + 1
					If objTest.Assertions(iAssertIndex).Passed Then
						objModule.PassCount = objModule.PassCount + 1
					End If
				Next

				Set objTest = Nothing
			Next
			intTimeEnd = Timer

			objModule.Duration = (intTimeEnd - intTimestart) * 1000

			Set m_CurrentModule = Nothing
		End Sub

		Private Sub RunModuleTest(ByRef objTest)
			Set m_CurrentTest = objTest

			On Error Resume Next
			Call GetRef(objTest.Name)()

			If Err.Number <> 0 Then
				Call Assert(False, Err.Source & " (Code: " & Err.Number & "), " & Err.Description & ".")
			End If

			Err.Clear()
			On Error Goto 0

			Set m_CurrentTest = Nothing
		End Sub

		Private Sub RunTestModuleSetup(ByRef objModule)
			If Not objModule.Lifecycle Is Nothing Then
				If Not objModule.Lifecycle.Setup = Empty Then
					Call GetRef(objModule.Lifecycle.Setup)()
				End If
			End If
		End Sub

		Private Sub RunTestModuleTeardown(ByRef objModule)
			If Not objModule.Lifecycle Is Nothing Then
				If Not objModule.Lifecycle.Teardown = Empty Then
					Call GetRef(objModule.Lifecycle.Teardown)()
				End If
			End If
		End Sub

		Private Function matchDictionaries(dic1, dic2)
			matchDictionaries = False
			If typeName(dic1) <> "Dictionary" or typeName(dic2) <> "Dictionary" Then _
				Exit Function
			If dic1.Count <> dic2.Count Then _
				Exit Function
			Dim dic1Keys : dic1Keys = dic1.Keys
			Dim dic2Keys : dic2Keys = dic2.Keys
			If UBound(dic1Keys) <> UBound(dic2Keys) Then _
				Exit Function
			Dim i, current
			For i = 0 To UBound(dic1Keys)
				current = dic1Keys(i)
				' For each key make sure a same key exists in the other dictionary.
				If Not dic2.Exists(current) Then _
					Exit Function
				' Make sure the typeName of the variables are equal.
				If Not typeName(dic1(current)) = typeName(dic2(current)) Then _
					Exit Function
				' Comparison of Nothing objects should be done with 'is', but redundant.
				If typeName(dic1(current)) <> "Nothing" Then
					' Sub-dictionary objects should be matched too by passing it into a sub-call.
					If typeName(dic1(current)) = "Dictionary" Then
						If Not matchDictionaries(dic1(current), dic2(current)) Then _
							Exit Function
					Else
						' For other types the content should match.
						If Not dic1(current) = dic2(current) Then _
							Exit Function
					End If
				End If
			Next
			matchDictionaries = True
		End Function
	End Class

	' Private Classses

	Class ASPUnitScenario
		Public _
			Modules, _
			TestCount, _
			PassCount, _
			FailCount

		Private Sub Class_Initialize()
			Set Modules = Server.CreateObject("System.Collections.ArrayList")
			PassCount = 0
			TestCount = 0
			FailCount = 0
		End Sub

		Private Sub Class_Terminate()
			Set Modules = Nothing
		End Sub

		Public Property Get Passed
			Passed = (PassCount = TestCount)
		End Property
	End Class

	Class ASPUnitModule
		Public _
			Name, _
			Tests, _
			Lifecycle, _
			Duration, _
			TestCount, _
			PassCount

		Private Sub Class_Initialize()
			Set Tests = Server.CreateObject("System.Collections.ArrayList")
			TestCount = 0
			PassCount = 0
		End Sub

		Private Sub Class_Terminate()
			Set Tests = Nothing
		End Sub

		Public Property Get FailCount
			FailCount = (TestCount - PassCount)
		End Property

		Public Property Get Passed
			Passed = (PassCount = TestCount)
		End Property
	End Class

	Class ASPUnitTest
		Public _
			Name, _
			Passed, _
			Assertions
		
		Private Sub Class_Initialize
			Set Assertions = Server.CreateObject("Scripting.Dictionary")
			Passed = True
		End Sub

		Private Sub Class_Terminate
			Set Assertions = Nothing
		End Sub
	End Class

	Class ASPUnitTestAssertion
		Public _
			Passed, _
			Description, _
			Actual, _
			Expected
		
		Private Sub Class_Initialize
			Actual = Empty
			Expected = Empty
		End Sub
	End Class

	Class ASPUnitTestLifecycle
		Public _
			Setup, _
			Teardown
	End Class
%>