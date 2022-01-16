;Written by KramWell.com - 22/JAN/2019
;Simple program to send a basic test email to a specified SMTP server over port 25.

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
 #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#NoTrayIcon

Gui,Add,Text,x10 y13 w40 h21,From:
Gui,Add,Edit,x50 y10 w120 h21 vfromAddress,

Gui,Add,Text,x10 y43 w40 h21,To:
Gui,Add,Edit,x50 y40 w120 h21 vtoAddress,

Gui,Add,Text,x10 y73 w40 h21,SMTP:
Gui,Add,Edit,x50 y70 w120 h21 vsmtpAddress,

Gui,Add,Button,x15 y100 w150 h23 gSendEmail,Send Email

Gui,Show,w190 h140,SMTP Test

Return


SendEmail:

	GuiControlGet, fromAddress
	GuiControlGet, toAddress
	GuiControlGet, smtpAddress

	If (fromAddress = "") OR (toAddress = "") OR (smtpAddress = ""){	
		MsgBox % "You have blank fields, please correct."
		Return ;return if error, continue if not
	}	

	;check to see if valid
	If !InStr(fromAddress, "@"){
		MsgBox % "The from address '" fromAddress "' is not valid."
		Return ;return if error, continue if not
	}			
	If !InStr(toAddress, "@"){
		MsgBox % "The to address '" toAddress "' is not valid."
		Return ;return if error, continue if not
	}	
	If !InStr(smtpAddress, "."){
		MsgBox % "The SMTP address '" smtpAddress "' is not valid."
		Return ;return if error, continue if not
	}	

	pmsg 						:= ComObjCreate("CDO.Message")
	pmsg.From 					:= fromAddress
	pmsg.To 					:= toAddress
	pmsg.Subject 					:= "test email"
	pmsg.TextBody 					:= "Testing SMTP from server " . A_ComputerName " from user " . A_UserName

	fields 						:= Object()
	fields.smtpserver   				:= smtpAddress
	fields.smtpserverport   			:= 25 										; 25
	fields.smtpusessl      				:= False 									; False
	fields.sendusing     				:= 2   										; cdoSendUsingPort
	fields.smtpauthenticate 			:= 0   										; cdoBasic
	fields.smtpconnectiontimeout			:= 10
	schema 						:= "http://schemas.microsoft.com/cdo/configuration/"
	pfld 						:=  pmsg.Configuration.Fields
	For field,value in fields
		pfld.Item(schema . field) 		:= value
	pfld.Update()

	try  ; Attempts to send email.
	{
		pmsg.Send()
	}
		catch e
		{
			MsgBox, 16,, % "Exception thrown!`n`nwhat: " e.what "`nfile: " e.file
				. "`nline: " e.line "`nmessage: " e.message "`nextra: " e.extra
			Return
		}

	MsgBox % "Email Sent!"

Return

GuiClose:
ExitApp 