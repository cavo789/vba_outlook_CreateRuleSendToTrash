' ==============================================================
' Description: Outlook macro for creating a rule that will send all
' emails received from a specific sender or domain to the trash.
'
' How to use (once installed) :
'	1. From Outlook, select one or more spam emails
'	2. Run the macro below (best : add a button in your ribbon)
'
' What does this macro ?
'	* Get the email address of the spammer (f.i. spam@yahoo.co.uk)
'	* Ask the user if emails from that domain (@yahoo.co.uk) should 
'		be removed from now or only those sender (spam@yahoo.co.uk)
'	* Create a rule in the "DefaultStore" of Outlook for removing
'		emails with that emails/domain and send mails to the deleted
'		items folder
'	* Iterate all accounts present in your Outlook and run rules
'		against all accounts. So, every if the rule is created only
'		in the "DefaultStore" (i.e. one account), the rule will be
'		fired for all configured accounts.
'
' Author : Christophe Avonture
' Inspired by @link : https://msdn.microsoft.com/en-us/vba/outlook-vba/articles/create-a-rule-to-move-specific-e-mails-to-a-folder and by
' https://social.msdn.microsoft.com/Forums/office/en-US/161ca567-9b6f-4155-9833-cc371d78b66e/vba-macro-to-run-rules-for-multiple-email-accounts?forum=outlookdev
'====================================================

Option Explicit
Option Base 0

Public Sub CreateRuleSendToTrash()

	' Prefix to use for naming rules created by this macro
	Const cPrefix = "SendToTrash - "

	Dim oRules As Outlook.Rules
	Dim oRule As Outlook.Rule
	Dim oMail As Outlook.MailItem
	Dim oExplorer As Outlook.Explorer
	Dim oSession As Outlook.NameSpace
	Dim oAccount As Outlook.Account
	Dim oDelivery As Outlook.Store
	Dim oRuleCondition As Object ' Use Late binding
	Dim oMoveRuleAction As Outlook.MoveOrCopyRuleAction
	Dim sSendereMail As String, sRuleName As String
	Dim bContinue As Boolean, bDomain As Boolean
	Dim I As Integer, J As Integer, K As Byte
	Dim arrTemp() As String
	Dim sDomain As String

	Set oExplorer = Application.ActiveExplorer

	' Make sure at least one item is selected
	If oExplorer.Selection.Count = 0 Then
		Call MsgBox("Please select at least one email before", vbExclamation, _
			"Create rule to send emails to trash")
		Exit Sub
	End If

	J = oExplorer.Selection.Count

	' Process every selected emails; one by one
	For I = 1 To J

		' Retrieve the selected email
		Set oMail = oExplorer.Selection.Item(I)

		' Mark the mail as Read
		oMail.UnRead = False

		' Retrieve the spam email address, the one to ban (f.i. spam@yahoo.jp)
		sSendereMail = oMail.SenderEmailAddress

		' Retrieve the domain name (f.i. @yahoo.jp)
		arrTemp = Split(sSendereMail, "@")
		sDomain = "@" + arrTemp(1)

		' --------------------------------------
		' 1. Verify if that rule doesn't exists yet

		' Define the name of the rule
		For K = 1 To 2

			sRuleName = cPrefix & IIf(K = 1, sSendereMail, sDomain)

			bContinue = True
			Set oRules = Session.DefaultStore.GetRules()
		
			For Each oRule In oRules
					If (oRule.Name = sRuleName) Then
						Debug.Print sRuleName & " already exists"
						K = 2
						bContinue = False
						Exit For
					End If
			Next

		Next K
	 
		If (bContinue) Then

			' ----------------------------------------
			' 2. The rule doesn't exists yet so create it

			' Do we need to block only the sender (spam@yahoo.jp) or all emails coming from the domain (@yahoo.jp)

			bDomain = (MsgBox("SendToTrash - Send emails to Trash as soon as they " & _
				"are received " & vbCrLf & vbCrLf & _
				"* from domain " & sDomain & " (--> Click on Yes) " & vbCrLf & vbCrLf & _
				"* from sender " & sSendereMail & " (--> Click on No) ", _
				vbQuestion + vbYesNo + vbDefaultButton1) = vbYes)

			Debug.Print "Create a rule for removing all mails sent by " + _
				IIf(bDomain, sDomain, sSendereMail)
		
			If (bDomain) Then
				'Set oRuleCondition = Outlook.AddressRuleCondition
				Set oRule = oRules.Create(cPrefix & sDomain, olRuleReceive)
			Else
				'Set oRuleCondition = Outlook.ToOrFromRuleCondition
				Set oRule = oRules.Create(cPrefix & sSendereMail, olRuleReceive)
			End If

			' Condition = when received by the spam email address
			If (bDomain) Then
				Set oRuleCondition = oRule.Conditions.SenderAddress
				oRuleCondition.Address = Array(sDomain)
			Else
				Set oRuleCondition = oRule.Conditions.From
				oRuleCondition.Recipients.Add (sSendereMail)
			End If

			With oRuleCondition
				.Enabled = True
				If Not (bDomain) Then .Recipients.ResolveAll
			End With

			' Action = send to the trash
			Set oMoveRuleAction = oRule.Actions.MoveToFolder
			With oMoveRuleAction
				.Enabled = True
				.Folder = Session.GetDefaultFolder(olFolderDeletedItems)
			End With

			' Save the rules collection
			oRules.Save

		End If ' If bContinue

	Next I

	' ----------------------------------------------------------
	' 3. Runs every rules on every accounts present in Outlook

	bContinue = (MsgBox("Run the rules now (YES) or wait until the next load of Outlook (NO)?", _
		vbQuestion + vbYesNo + vbDefaultButton1) = vbYes)

	If bContinue Then

		Set oSession = Application.Session

		' Iterate all accounts
		For Each oAccount In oSession.Accounts

			' For each account that is Exchange or iMap, try to run rules
			If oAccount.AccountType = olExchange Or oAccount.AccountType = olImap Then

				On Error Resume Next
				bContinue = Not (oAccount.DeliveryStore Is Nothing)
				If Err.Number <> 0 Then
					bContinue = False
					Err.Clear
				End If
				On Error GoTo 0

				If bContinue Then
					' There are rules to process
					Debug.Print "Processing rules for account: " & _
						oAccount.DisplayName & vbCrLf & _
						"===========================" & vbCrLf

					' Get the store for this account
					Set oDelivery = oAccount.DeliveryStore

					' Now iterate through the rules
					' DON'T USE "Set oRules = oDelivery.GetRules" since
					' this macro will add every rules in one account for
					' bigger simplification (don't need to search a rule
					' in the 10 configured accounts f.i.).
					' So, oRules is the "Session.DefaultStore.GetRules()"
					' Get the list of rules from there

					For Each oRule In oRules
						If oRule.RuleType = olRuleReceive And oRule.Enabled Then
							If Left(oRule.Name, 14) = cPrefix Then
								Debug.Print oRule.Name
								oRule.Execute RuleExecuteOption:=OlRuleExecuteOption.olRuleExecuteAllMessages, _
									Folder:=oDelivery.GetDefaultFolder(olFolderInbox)
							End If
						End If
					Next ' oRule
				End If ' If bContinue Then

			End If ' If oAccount.AccountType

		Next ' For Each oAccount

	End If ' If bContinue Then

End Sub
