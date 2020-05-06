<% Option Explicit %>

<!-- #include file="../Lib/ASPUnit.asp" -->

<%
	' Register pages to test
	Call ASPUnit.AddPages(Array( _
		"ASPUnitLibrary.asp", _
		"ASPUnitTesterFactoryMethods.asp", _
		"ASPUnitTesterAssertionMethods.asp", _
		"ASPUnitTesterControlMethods.asp", _
		"ASPUnitTesterBehaviors.asp", _
		"ASPUnitRunner.asp" _
	))

	' Execute tests
	Call ASPUnit.Run()
%>