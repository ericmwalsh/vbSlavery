'Kick myself out of a loan
Option Explicit
'2014-02-12 AUTHOR: Eric M Walsh
Sub Main
	Call kickMvLoan
End Sub

Public Sub kickMvLoan

	Dim current_user As String
	Dim loan_id As String

	'This portion sets the current_user variable to the current Morvision username (e.g. ERIC.WALSH) using the get_current_user.bas script.
	'#uses "get_current_user.bas"
	current_user = get_current_user

	'Sets the loan_id variable to the Loan-ID using the GetCalcString API Method
	loan_id = GetCalcString(7073)

	Dim str_sql As String
	
	const ui_title = "Kick myself out of loan (hotkick_out_of_loan.bas)"
	

	'#uses "database_connection.bas"
	open_database_connection

		str_sql = "SELECT * FROM TABLENAME WHERE User_ID = '" & current_user & "' AND Loan_ID != '" & loan_id & "' ORDER BY Loan_Lock_Date asc"		
		Set adoMRS = adoMCN.Execute (str_sql)

		If adoMRS.EOF then
			MsgBox "ERROR: The current user(" & current_user & ") is not logged into any loans except this one!", vbcritical, ui_title
		Else
			Dim temp_loan_id As String
			Do
				temp_loan_id = adoMRS.Fields(0)
				If MsgBox ("Would you like to be kicked out of Loan: " & temp_loan_id & " ?", vbYesNo, ui_title) = vbYes Then
					str_sql = "DELETE FROM TABLENAME WHERE User_ID = '" & current_user & "' AND Loan_ID = '" & temp_loan_id & "'"
					adoMCN.Execute(str_sql)
					MsgBox "User " & current_user & " was kicked out of a loan " & temp_loan_id & ".", vbcritical, ui_title
				End If
				adoMRS.MoveNext
			Loop Until adoMRS.EOF
			
		End If

	close_database_connection

End Sub