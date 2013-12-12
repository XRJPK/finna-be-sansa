#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Compression=4
#AutoIt3Wrapper_Res_Fileversion=0.1.1.4
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=p
#AutoIt3Wrapper_Res_LegalCopyright=© Th. Meyer 2013
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiStatusBar.au3>
#include <WindowsConstants.au3>
#include <String.au3>
#Region ### START Koda GUI section ### Form=F:\Sotec\_SCRIPTS\AutoIT\GUI\Forms\MSI.kxf
$MSI = GUICreate("Simple MSI Editor", 598, 438, 192, 124)
$MenuItem1 = GUICtrlCreateMenu("Datei")
$Beenden = GUICtrlCreateMenuItem("Beenden", $MenuItem1)
$MenuItem2 = GUICtrlCreateMenu("?")
$StatusBar1 = _GUICtrlStatusBar_Create($MSI)
_GUICtrlStatusBar_SetMinHeight($StatusBar1, 20)
$Button1 = GUICtrlCreateButton("MSI öffnen", 24, 32, 80, 25)
$Button2 = GUICtrlCreateButton("MSI schreiben", 24,60, 80, 25)
$Button3 = GUICtrlCreateButton("Reverse GUID", 24,88, 80, 25)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

;~ #####################################################################################
;~ #                                                                                   #
;~ #					            	CHANGELOG                                       #
;~ #                                                                                   #
;~ #####################################################################################

;~ 0.1.1.0
;~ 	29.10.2013 - by Th. Meyer
;~ 	Initial Release of Version 0.1.1.0
;~  0.1.1.4
;~ 	29.10.2013 - by Th. Meyer
;~ 	Added generateGUID Function

;~  0.1.1.5
;~ 	30.10.2013 - by Th. Meyer
;~  added TabControl

;~ ########################### END CHANGELOG ###########################################

Local $verzeichnis= ""
Local $msi_array[2]
Local $record[99]
Local $MSI = ObjCreate("WindowsInstaller.Installer")
Local $DB = ""
Local $View = ""
Local $Summary = ""
Local $GUIDString

;                                     DOKUMENTATION usw .....
; #########################################################################################
;
; ============ MSI Summary Zuordnung ==============
    ;  $Summary.Property(1) ; SI
	;  $Summary.Property(2); Title
	;  $Summary.Property(3) ; Subject
	;  $Summary.Property(4) ; author
	;  $Summary.Property(5) ; keywords
	;  $Summary.Property(6) ; comments
	;  $Summary.Property(7) ; platform
	;  $Summary.Property(8) ; last seved
	;  $Summary.Property(9) ; Package code
	;  $Summary.Property(10) ;
	;  $Summary.Property(11) ;
	;  $Summary.Property(12) ;
	;  $Summary.Property(13) ;
	;  $Summary.Property(14) ; Schema
	;  $Summary.Property(15) ;
	;  $Summary.Property(16) ;
	;  $Summary.Property(17) ;
	;  $Summary.Property(18) ; Creating App
	;  $Summary.Property(19) ;
; ======= END - MSI Summary Zuordnung ==============


While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
	Case $GUI_EVENT_CLOSE
			Exit

	Case $MenuItem2
			MsgBox (0, "Über", "MSI Toolbox c Th. Meyer & Fr. Minklei 2013", 10)
	Case $Button1 ; MSI öffnen
			;==========================================================================================================================
			;  MSI Öffnen Dialog, es werden nur MSIs angezeigt und anschließend wird der Pfad in die $verzeichnis Variable gespeichert!
			;==========================================================================================================================
			$verzeichnis = FileOpenDialog ("Bitte MSI wählen", $verzeichnis, "Installer Files (*.msi;*.wsi) | All Files (*.*)")
			;MsgBox(0, "Debug", "Das Verzechniss ist: "  & $verzeichnis, 10)

			$MSI     = ObjCreate("WindowsInstaller.Installer")
			$Summary = $MSI.SummaryInformation($verzeichnis,0)
			$Tab1 = GUICtrlCreateTab(320, 10, 270, 350, 1)
			$TabSheet1 = GUICtrlCreateTabItem("Summary Information" & $Summary.Property(2))
			$Title = GUICtrlCreateLabel("Title", 325, 40, 24, 17)
			$Author = GUICtrlCreateLabel("Author", 325, 64, 35, 17)
			$Plattform = GUICtrlCreateLabel("Plattform", 325, 88, 45, 17)
			$Langauges = GUICtrlCreateLabel("Langauges", 325, 112, 57, 17)
			$Package_Code = GUICtrlCreateLabel("Package Code", 325, 136, 75, 17)
			$Schema = GUICtrlCreateLabel("Schema", 325, 160, 43, 17)
			$Input1 = GUICtrlCreateInput("", 410, 40, 170, 21)
			GUICtrlSetState(-1, $GUI_DISABLE)
			$Input2 = GUICtrlCreateInput("", 410, 64, 170, 21)
			GUICtrlSetState(-1, $GUI_DISABLE)
			$Input3 = GUICtrlCreateInput("", 410, 88, 170, 21)
			GUICtrlSetState(-1, $GUI_DISABLE)
			$Input4 = GUICtrlCreateInput("", 410, 112, 170, 21)
			GUICtrlSetState(-1, $GUI_DISABLE)
			$Input5 = GUICtrlCreateInput("", 410, 136, 170, 21)
			$Input6 = GUICtrlCreateInput("", 410, 160, 170, 21)
			$Input7 = GUICtrlCreateInput("GUID", 330, 188, 240, 21)
			GUICtrlSetState(-1, $GUI_DISABLE)
			$Input8 = GUICtrlCreateInput("Reverse GUID", 330, 222, 240, 21)
			;GUICtrlCreateGroup("", -99, -99, 1, 1)
			;GUICtrlCreateGroup("", -99, -99, 1, 1)
			GUICtrlCreateTabItem("")

		;	$Group1 = GUICtrlCreateGroup("Summary Information von " & $Summary.Property(2),  328, 8, 257, 361)

			GuiCtrlSetData( $Input1,  $Summary.Property(2))
			GuiCtrlSetData( $Input2,  $Summary.Property(4))
			GuiCtrlSetData( $Input3,  $Summary.Property(7))
			GuiCtrlSetData( $Input4,  $Summary.Property(7))
			GuiCtrlSetData( $Input5,  $Summary.Property(9))
			GuiCtrlSetData( $Input6,  $Summary.Property(14))

			$DB      = $MSI.OpenDataBase ($verzeichnis,1)

		;$View    = $DB.OpenView("SELECT 'Value') FROM 'Property' WHERE 'Property' = ProductCode")
	Case $Button2 ; MSI schreiben
;~ 			MsgBox (0, "Debug", $DB.LastErrorRecord())
;~ 			$View    = $DB.OpenView("SELECT `Value` FROM `Property` WHERE `Property` = 'ARPCONTACT'")
;~ 			MsgBox (0, "Debug", $DB.LastErrorRecord())
;~ 			$View.Execute
;~ 			$record= $View.Fetch
;~ 			$View.Close
			$View    = $DB.OpenView("SELECT `Value` FROM `Property` WHERE `Property` = 'ARPCONTACT'")
			$View.Execute
			$record= $View.Fetch
			$record = $View.Modify

	Case $Button3
		$View    = $DB.OpenView("SELECT `Value` FROM `Property` WHERE `Property` = 'ProductCode'")
	    $View.Execute
	    $record= $View.Fetch
		GuiCtrlSetData( $Input7,  $record.StringData(1))
		GUICtrlSetData ( $Input8, generateGUID())

	EndSwitch
WEnd

; ============================================================
;  Functions Section
; ============================================================

Func reverseString( $str )
$reverseString = _StringReverse( StringMid($str,1,8) ) _
		& _StringReverse( StringMid($str,9,4) ) _
		& _StringReverse( StringMid($str,13,4) ) _
		& _StringReverse( StringMid($str,17,2) ) _
		& _StringReverse( StringMid($str,19,2) ) _
		& _StringReverse( StringMid($str,21,2) ) _
		& _StringReverse( StringMid($str,23,2) ) _
		& _StringReverse( StringMid($str,25,2) ) _
		& _StringReverse( StringMid($str,27,2) ) _
		& _StringReverse( StringMid($str,29,2) ) _
		& _StringReverse( StringMid($str,31,2) )
	EndFunc


Func generateGUID ()
	$typeLib = ObjCreate("Scriptlet.TypeLib")
	Return $typeLib.Guid
	EndFunc
