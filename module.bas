' ==============================================================
' Description: Outlook macro for creating a rule that will send all
' 	emails received from a specific sender to the trash.
'
' How to use (once installed) :
'	1. From Outlook, select the mail of the spammer
'	2. Run the macro below (best : add a button in your ribbon)
'
' What does this macro ?
'	* Get the email address of the spammer (f.i. contact@bank.com)
'	* Create a rule in the "DefaultStore" of Outlook for removing
'		emails with that emails and send mails to the deleted items
'		folder
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
	Dim oFromCondition As Outlook.ToOrFromRuleCondition
	Dim oMoveRuleAction As Outlook.MoveOrCopyRuleAction
	Dim sSendereMail As String, sRuleName As String
	Dim bContinue As Boolean

	Set oExplorer = Application.ActiveExplorer

	' Make sure at least one item is selected
	If oExplorer.Selection.Count <> 1 Then
		Call MsgBox("Please select a single item", vbExclamation, _
			"Create rule to send emails to trash")
		Exit Sub
	End If

	' Retrieve the selected email
	Set oMail = oExplorer.Selection.Item(1)

	' Retrieve the spam email address, the one to ban
	sSendereMail = oMail.SenderEmailAddress

	' --------------------------------------
	' 1. Verify if that rule doesn't exists yet

	' Define the name of the rule
	sRuleName = cPrefix & sSendereMail

	bContinue = True
	Set oRules = Session.DefaultStore.GetRules()

	For Each oRule In oRules
		If (oRule.Name = sRuleName) Then
			bContinue = False
			Exit For
		End If
	Next

	If (bContinue) Then

		' ----------------------------------------
		' 2. The rule doesn't exists yet so create it

		Debug.Print "Create a rule for removing all mails sent by " + _
			sSendereMail

		Set oRule = oRules.Create(cPrefix & sSendereMail, olRuleReceive)

		' Condition = when received by the spam email address
		Set oFromCondition = oRule.Conditions.From
		With oFromCondition
			.Enabled = True
			.Recipients.Add (sSendereMail)
			.Recipients.ResolveAll
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

	' ----------------------------------------------------------
	' 3. Runs every rules on every accounts present in Outlook

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
			End If
		End If
	Next ' oAccount

End Sub
