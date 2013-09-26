<%
'********************************************************************
' Name: ASPUnitRunner.asp
'
' Purpose: Contains the UnitRunner class which is used to render the unit testing UI
'********************************************************************

'********************************************************************
' Include Files
'********************************************************************
%>
<!-- #include file="ASPUnit.asp"-->
<%

Const ALL_TESTCONTAINERS = "所有测试类"
Const ALL_TESTCASES = "所有测试案例"

Const FRAME_PARAMETER = "UnitRunner"
Const FRAME_SELECTOR = "selector"
Const FRAME_RESULTS = "results"

Const STYLESHEET = "include/ASPUnit.css"
Const SCRIPTFILE = "include/ASPUnitRunner.js"

Class UnitRunner

	Private m_dicTestContainer

	Private Sub Class_Initialize()
		Set m_dicTestContainer = CreateObject("Scripting.Dictionary")
	End Sub

	Public Sub AddTestContainer(oTestContainer)
		m_dicTestContainer.Add TypeName(oTestContainer), oTestContainer
	End Sub

	Public Function Display()
		If (Request.QueryString(FRAME_PARAMETER) = FRAME_SELECTOR) Then
			DisplaySelector
		ElseIf (Request.QueryString(FRAME_PARAMETER) = FRAME_RESULTS) Then
			DisplayResults
		Else
			ShowFrameSet
		End if
	End Function

'********************************************************************
' Frameset
'********************************************************************
	Private Function ShowFrameSet()
%>
<HTML>
<HEAD>
<TITLE>ASPUnit Test Runner</TITLE>
</HEAD>
<FRAMESET ROWS="30, *" BORDER=0 FRAMEBORDER=0 FRAMESPACING=0>
	<FRAME NAME="<% = FRAME_SELECTOR %>" src="<% = GetSelectorFrameSrc %>" marginwidth="0" marginheight="0" scrolling="auto" border=0 frameborder=0 noresize>
	<FRAME NAME="<% = FRAME_RESULTS %>" src="<% = GetResultsFrameSrc %>" marginwidth="0" marginheight="0" scrolling="auto" border=0 frameborder=0 noresize>
</FRAMESET>
<%
	End Function

	Private Function GetSelectorFrameSrc()
		GetSelectorFrameSrc = Request.ServerVariables("SCRIPT_NAME") & "?" & FRAME_PARAMETER & "=" & FRAME_SELECTOR
	End Function

	Private Function GetResultsFrameSrc()
		GetResultsFrameSrc = Request.ServerVariables("SCRIPT_NAME") & "?" & FRAME_PARAMETER & "=" & FRAME_RESULTS
	End Function

'********************************************************************
' Selector Frame
'********************************************************************
	Private Function DisplaySelector()
%>
<HTML>
<HEAD>
<LINK REL="stylesheet" HREF="<% = STYLESHEET %>" MEDIA="screen" TYPE="text/css">
<SCRIPT>
function ComboBoxUpdate(strSelectorFrameSrc, strSelectorFrameName)
{
	document.frmSelector.action = strSelectorFrameSrc;
	document.frmSelector.target = strSelectorFrameName;
	document.frmSelector.submit();
}
</SCRIPT>
</HEAD>
<BODY>
		<FORM NAME="frmSelector" ACTION="<% = GetResultsFrameSrc %>" TARGET="<% = FRAME_RESULTS %>" METHOD=POST>
			<TABLE WIDTH="80%">
				<TR>
					<TD ALIGN="right">测试:</TD>
					<TD>
						<SELECT NAME="cboTestContainers" OnChange="ComboBoxUpdate('<% = GetSelectorFrameSrc %>', '<% = FRAME_SELECTOR %>');">
						<OPTION><% = ALL_TESTCONTAINERS %>
<%
							AddTestContainers
%>
						</SELECT>
					</TD>
					<TD ALIGN="right">测试方法:</TD>
					<TD>
						<SELECT NAME="cboTestCases">
						<OPTION><% = ALL_TESTCASES %>
<%
							AddTestMethods
%>
						</SELECT>
					<TD>
						<INPUT TYPE="checkbox" NAME="chkShowSuccess"> 显示通过的测试</INPUT>
					</TD>
					</TD>
					<TD>
						<INPUT TYPE="Submit" NAME="cmdRun" VALUE="运行测试">
					</TD>
				</TR>
			</TABLE>
		</FORM>
</BODY>
</HTML>
<%
	End Function

	Private Function AddTestContainers()
		Dim oTestContainer, sTestContainer
		For Each oTestContainer In m_dicTestContainer.Items()
			sTestContainer = TypeName(oTestContainer)
			If (sTestContainer = Request.Form("cboTestContainers")) Then
				Response.Write("<OPTION SELECTED>" & sTestContainer)
			Else
				Response.Write("<OPTION>" & sTestContainer)
			End If
		Next
	End Function

	Private Function AddTestMethods()
		Dim oTestContainer, sContainer, sTestMethod

		If (Request.Form("cboTestContainers") <> ALL_TESTCONTAINERS And Request.Form("cboTestContainers") <> "") Then
			sContainer = CStr(Request.Form("cboTestContainers"))
			Set oTestContainer = m_dicTestContainer.Item(sContainer)
			For Each sTestMethod In oTestContainer.TestCaseNames()
				Response.Write("<OPTION>" & sTestMethod)
			Next
		End If
	End Function

	Private Function TestName(oResult)
		If (oResult.TestCase Is Nothing) Then
			TestName = ""
		Else
			TestName = TypeName(oResult.TestCase.TestContainer) & "." & oResult.TestCase.TestMethod
		End If
	End Function

'********************************************************************
' Results Frame
'********************************************************************
	Private Function DisplayResults()
%>
<HTML>
<HEAD>
<LINK REL="stylesheet" HREF="<% = STYLESHEET %>" MEDIA="screen" TYPE="text/css">
</HEAD>
<BODY>
<%
		Dim oTestResult, oTestSuite
		Set oTestResult = New TestResult

		' Create TestSuite
		Set oTestSuite = BuildTestSuite()

		' Run Tests
		oTestSuite.Run oTestResult

		' Display Results
		DisplayResultsTable oTestResult
%>
</BODY>
</HTML>
<%
	End Function

	Private Function BuildTestSuite()

		Dim oTestSuite, oTestContainer, sContainer
		Set oTestSuite = New TestSuite

		If (Request.Form("cmdRun") <> "") Then
			If (Request.Form("cboTestContainers") = ALL_TESTCONTAINERS) Then
				For Each oTestContainer In m_dicTestContainer.Items()
					If Not(oTestContainer Is Nothing) Then
						oTestSuite.AddAllTestCases oTestContainer
					End If
				Next
			Else
				sContainer = CStr(Request.Form("cboTestContainers"))
				Set oTestContainer = m_dicTestContainer.Item(sContainer)

				Dim sTestMethod
				sTestMethod = Request.Form("cboTestCases")

				If (sTestMethod = ALL_TESTCASES) Then
					oTestSuite.AddAllTestCases oTestContainer
				Else
					oTestSuite.AddTestCase oTestContainer, sTestMethod
				End If
			End If
		End If

		Set BuildTestSuite = oTestSuite
	End Function

	Private Function DisplayResultsTable(oTestResult)
%>
			<TABLE BORDER="1" WIDTH="80%">
				<TR><TH WIDTH="10%">类型</TH><TH WIDTH="20%">测试项</TH><TH WIDTH="70%">描述</TH></TR>
<%
		If Not(oTestResult Is Nothing) Then
			Dim oResult
			If (Request.Form("chkShowSuccess") <> "") Then
	            For Each oResult in oTestResult.Successes
					Response.Write("	<TR CLASS=""success""><TD>通过</TD><TD>" & TestName(oResult) & "</TD><TD>" & oResult.Source & oResult.Description & "</TD></TR>")
	            Next
	        End If

			For Each oResult In oTestResult.Errors
				Response.Write("	<TR CLASS=""error""><TD>错误</TD><TD>" & TestName(oResult) & "</TD><TD>" & oResult.Source & " (" & Trim(oResult.ErrNumber) & "): " & oResult.Description & "</TD></TR>")
			Next

			For Each oResult In oTestResult.Failures
				Response.Write("	<TR CLASS=""warning""><TD>故障</TD><TD>" & TestName(oResult) & "</TD><TD>" & oResult.Description & "</TD></TR>")
			Next

			Response.Write("	<TR><TD ALIGN=""center"" COLSPAN=3>" & "测试个数: " & oTestResult.RunTests & ", 错误个数: " & oTestResult.Errors.Count & ", 故障个数: " & oTestResult.Failures.Count & "</TD></TR>")
		End If
%>
			</TABLE>
<%
	End Function

	Public Sub OnStartTest()

	End Sub

	Public Sub OnEndTest()

	End Sub

	Public Sub OnError()

	End Sub

	Public Sub OnFailure()

	End Sub

    Public Sub OnSuccess()

    End Sub
End Class
%>

