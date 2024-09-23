/********************************
* NoteAssist V2.2 - 2024-JUL-01 *
* Written by John Campbell	    *
* Intended for use with AHK V2  *
********************************/

#SingleInstance
#NoTrayIcon
#Include UIA.ahk
#HotIf WinActive("NoteAssist")
^Backspace::SendInput "+{Left}" "{Delete}"
guititle := "NoteAssist - *TESTING ENVIRONMENT*"
Global Version := "*TESTING ENVIRONMENT*"
checkmark := "X"
checkdate := A_now
checkdate := FormatTime(checkdate,"MMdd")
If checkdate = "1031"
checkmark := "💀"
If checkdate = "1225"
checkmark := "🎅"
latexe := "Latitude - Account Work Form"

;MsgBox "This is the testing environment for NoteAssist. Please note that while the features added have been tested for accuracy and durability, that issues may still arise. If you encounter any unexpected behavior, please note the steps that you took prior to the issue and send a message to John."


	
/***********************************************
* Passes parameters to the compiler to         *
* to change the meta-data of the compiled file *
* This needs to remain commented out to work   *
***********************************************/

;@Ahk2Exe-SetProductName NoteAssist
;@Ahk2Exe-SetFileVersion 2.2.0.0
;@Ahk2Exe-SetVersion 2.2.0.0
;@Ahk2Exe-SetDescription NoteAssist
;@Ahk2Exe-SetCompanyName PFC

/********************************************
* Enables the program to call a windows DLL *
* This DLL allows us to scale with the DPI  *
* on non-primary monitors                   *
********************************************/

; https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setthreaddpiawarenesscontext

DllCall("SetThreadDpiAwarenessContext", "ptr", -3, "ptr")

	/*
	**
	*/

noteRegRead()

noteRegRead(){

/****************************************************
* This function reads the current_user_registry and *
* assigns the keys to values that are in our map    *
* this way our settings can be recalled			    *
*****************************************************/ 

try{
global settings := Map(
	"name"		,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "name"),
	"size"		,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "size"),
	"color"		,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "color"),
	"fontcolor"	,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "fontcolor"),
	"strike"	,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "strike"),
	"underline"	,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "underline"),
	"italic"	,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "italic"),
	"bold"		,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "bold"),
	"cc"		,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "cc"),
	"bgcolor"	,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "bgcolor"),
	"str"		,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "str"),
	"delim"		,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "delim"),
	"datefrmt"	,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "datefrmt"),
	"rbf"		,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "rbf"),
	"esls"		,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "esls"),
	"aaot"		,RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "aaot"),
)
}

IF A_LastError = 2{
try
if RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "rbf") = 1
MsgBox "There was an error loading your settings, they are being reset.","Error",48+262144
global settings := Map(
	"name"		,""	,
	"size"		,"10"				,
	"color"		,""					,
	"fontcolor"	,000000				,
	"strike"	,1					,
	"underline"	,1					,
	"italic"	,1					,
	"bold"		,1					,
	"cc"		,""					,
	"bgcolor"	,"F5F5F5"			,
	"str"		,"norm s10.0"		,
	"delim"		," - "				,
	"datefrmt"	,"MM-dd-yyyy"		,
	"rbf"		,1					,
	"esls"		,0					,
	"aaot"		,0					,
)
for header, val in settings
RegWrite val, "REG_SZ", "HKEY_CURRENT_USER\SOFTWARE\NoteAssist", header
}
}

NoteAssistGUI()

NoteAssistGUI(){

	;LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
	Global fontsize := settings["size"]
	Global noteGui 	:= gui(,guititle) ;"-DPIScale"
	Global tday     := A_Now
	Global dfrmt    := settings["datefrmt"]
	Global tday     := FormatTime(tday, dfrmt)
	Global delim    := settings["delim"]
	Global Margin   := fontsize
	noteGui.MarginX := noteGui.MarginY := Margin
	Global nttw    	:= Margin*8
	Global ntth    	:= Margin*2
	Global ntew    	:= Margin*20
	Global nteh    	:= Margin*2
	Global btnw   	:= Margin*9
	Global btnh   	:= Margin*4
	Global ntw     	:= Margin*2
	Global nth     	:= Margin*2
	
/************
*	MAPS	*
************/
	
	Global notectrl := Map( ;control map for text and edit controls
	"1","NAME:",
	"2","ADDRESS:",
	"3","PHONE:",
	"4","EMAIL:",
	"5","ADV BAL:",
	"6","REQ PIF:",
	"7","CALL NOTES:",
	)
	
	Global noteBMap := Map( ;button control map
	"1",	"Clear Notes",
	"2",	"Shortcuts",
	"3",	"Copy Notes",
	"4",	"Inbound",
	"5",	"Outbound",
	"6",	"Manual",
	"7",	"EPA",
	"8",	"F/C",
	"9",	"Multi Channel",
	; "10",	"",
	; "11",	"",
	)
	
	Global noteBSCMap := Map( ;button control map
	"001",	"Addresses"		,
	"002",	"Payment Matrix",
	"003",	"MM"			,
	"004",	"PIFLR"			,
	"005",	"SLEPA"			,
	"006",	"Dispute"		,
	"007",	"Consent  Email",
	"008",	"Phone Numbers" ,
	"009",	"Restrictions"	,
	"010",	"NA"			,
	"011",	"PUHU"			,
	"012",	"RTV"			,
	"013",	"Search Phone#"	,
	"014",	"MCAID"			,
	"015",	"HLD30"			,
	"016",	"NML"			,
	"017",	"WN"			,
	"018",	"TRC"			,
	"019",	"AML"			,
	"020",	"WORKING PQ"	,
	"021",	"ATTORNEY",
	

	)	
	
	Global noteCBm := Map( ;checkbox control map
	"cb1DOB"  	,"DOB"		  ,
	"cb2SSN" 	,"SSN"		  ,
	"cb3MM" 	,"MM"		  ,
	"cb4ls1"  	,"ADV INT TO REQ JMT UNLESS PIF" ,
	"cb5ls2"  	,"ADV WILL NOT ACT ON JUDGMENT UNLESS DEFAULTS ON PAYMENTS" ,
	)
	
	Global epacMap := Map( ;EPA Control Map
	"1",	"PAYOR NAME:"	,
	"2",	"ACCOUNT #:" 	,
	"3",	"BIN:"			,
	"4",	"LAST 4:"		,
	"5",	"BANK NAME:"	,
	"6",	"AMOUNT:"		,	
	"7",	"RECURRING?"	,
	"8",	"EPA:"			,
	"9",	"OK TO CB?"		,
	;"10",	""				,
	;"11",	""				,
	)

	Global fccMap := Map( ;F/C Control Map
	"1",	"Married/Single?:"	,
	"2",	"Rent/Own:" 		,
	"3",	"Employed:"			,
	"4",	"Bank Name:"		,
	)
	
	Global fccbMap := Map(
	"cb1ms",	"Single"		,
	"cb2ms",	"Married"		,
	"cb3ro",	"Rent"			,
	"cb4ro",	"Own"			,
	"cb5em",	"Unemployed"	,
	"cb6em",	"Employed"		,
	"cb7bn",	"Refused"		,
	"cb8bn",	"Consent"		,
	)

	Global slepaMap := Map( ;SLEPA Control Map
	"1",	"FULL NAME:"		,
	"2",	"LAST FOUR SSN:" 	,
	"3",	"EMAIL:"			,
	)	
	

/****************
*  GUI Options	*
****************/
DetectHiddenWindows true
noteGui.SetFont(settings["str"] " s" Margin " c" settings["fontcolor"], settings["name"])
if settings["aaot"] = 1
notegui.opt("AlwaysOnTop")
else
notegui.opt("-AlwaysOnTop")

/*********
* Header *
**********/
	if settings["esls"] = 0
	noteGui.Add("Text","xm w" ntew " h" Margin*3 " vTitle","Note Assist E/S")					;Title to fill empty space
	if settings["esls"] = 1
	noteGui.Add("Text","xm w" ntew " h" Margin*3 " vTitle","Note Assist L/S")					;Title to fill empty space
	noteGui["Title"].GetPos(&titleX, &titleY, &titleW, &titleH)
	noteGui.Add("Edit","xp+" titleW " yp+" margin/2 " w" nttw " 0x200 Center ReadOnly","" tday)	;Date Box
	noteGui["Title"].SetFont("S" Margin*1.75 " w" Margin*100)									;Set title opts
	
/************
*	MAIN	*
************/
	
		
		For con_name, header in notectrl{	;make text and edits using loop
				if A_Index < 7{
					noteGui.Add("Text","xm w" nttw " h" ntth " 0x200 vt" con_name, header)
					noteGui.Add("Edit","xp+" nttw " w" ntew " h" nteh " Uppercase 0x200 ve" con_name,"")
				} ;End If

				Else{ ;Makes the call notes
					;noteGui["t6"].GetPos(&tempx, &tempy, &tempw, &temph)
					noteGui.Add("Text","xm w" nteh*0.75 " h" nteh*0.75 " 0x200 vcbdisp Center BackgroundWhite border","")
					noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtdisp","Disputed")
					noteGui.Add("Text","yp w" nteh*0.75 " h" nteh*0.75 " 0x200 vcbcease Center BackgroundWhite border","")
					noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtcease","Cease")					
					noteGui.Add("Text","xp+" margin*6 " yp w" nteh*0.75 " h" nteh*0.75 " 0x200 vcbbad Center BackgroundWhite border","")
					noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtbad","Bad")
					noteGui.Add("Text","yp w" nteh*0.75 " h" nteh*0.75 " 0x200 vcbconsent Center BackgroundWhite border","")
					noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtconsent","Consent")					
					
					
					
					noteGui.Add("Text","xm w" nttw+ntew " h" ntth " Center Border 0x200 vt" con_name, header)
					noteGui.Add("Edit","yp+" ntth+(Margin/2) " w" ntew+nttw " h" nteh*5 " 0x200 Uppercase ve" con_name,"")
				} ;End Else
				If A_Index = 2{ ;Makes the CheckBoxes
					noteGui.Add("Text","xm+" nttw " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb1DOB Center BackgroundWhite border","")
					noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtDOB","DOB")
					noteGui.Add("Text","xp+" nttw/1.3 " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb2SSN Center BackgroundWhite border","")
					noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtSSN","SSN")
					noteGui.Add("Text","xp+" nttw/1.255 " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb3MM Center BackgroundWhite border","")
					noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtMM","MM")
					
					if settings["esls"] = 0
					settings["esls"] := 0
					
					if settings["esls"] = 1{
					noteGui.Add("Text","xm+" nttw " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb4ls1 Center BackgroundWhite border","")
					noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtls1","INT TO REQ JMT")
					noteGui.Add("Text","xp+" nttw+Margin*5.3 " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb5ls2 Center BackgroundWhite border","")
					noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtls2","WNA")
					noteGui.Add("Text","xm w" nttw+margin*5 " h" ntth " 0x200 vtls3","ADV BAL W/ FEES:")
					noteGui.Add("Edit","xp+" nttw+margin*5 " w" ntew/1.33 " h" nteh " Uppercase 0x200 vels3","")
					}
					
				} ;End CB If
		} ;End Loop
		

		
/*****************************
* 		 Button loops		 *
*****************************/

	/*************
	*    MAIN    *
	*************/
					
		new_row := 3		
		For bn, label in noteBMap
			if Mod(A_Index-1, new_row)
				con := noteGui.AddButton("x+" Margin/2 " w" btnw " h" btnh " vb" bn, "" label)
				,con.OnEvent("Click", bclick%bn%)
			else
				con := noteGui.AddButton("xm y+" Margin/2 " w" btnw " h" btnh " vb" bn, "" label)
				,con.OnEvent("Click", bclick%bn%)		

	/*************
	*     SC     *
	*************/
		noteGui["Title"].GetPos(&titleX, &titleY, &titleW, &titleH)		
		new_row := 3
		For bn, label in noteBSCMap
		if A_Index = 1{
		con := noteGui.AddButton("x" titleX+nttw+ntew+Margin " y" titleY " w" btnw " h" btnh " vsb" bn, "" label)
		,con.OnEvent("Click", sbclick%bn%)
		noteGui["sb001"].GetPos(&bax, &bay, &baw, &bah)
		}
		Else		
			if Mod(A_Index-1, new_row)
				con := noteGui.AddButton("x+" Margin/2 " w" btnw " h" btnh " vsb" bn, "" label)
				,con.OnEvent("Click", sbclick%bn%)
			else
				con := noteGui.AddButton("x" bax " y+" Margin/2 " w" btnw " h" btnh " vsb" bn, "" label)
				,con.OnEvent("Click", sbclick%bn%)
		noteGui["sb001"].GetPos(&bax, &bay, &baw, &bah)
		For bn, label in noteBSCMap{
		noteGui["sb" bn].Visible := 0		
		}

		;button move
		
		noteGui.Add("GroupBox", "x" titleX+nttw+ntew+Margin " y" titleY " w" btnw*3+(Margin/2)*4 " h" btnh*3-Margin " vgb1", "Information")
		
		noteGui["gb1"].GetPos(&gbax, &gbay, &gbaw, &gbah)
		noteGui["sb001"].GetPos(&bax1, &bay, &baw, &bah)
		noteGui["sb001"].Move(bax1+margin/2+1,gbay+Margin*2)
		noteGui["sb002"].GetPos(&bax2, &bay, &baw, &bah)
		noteGui["sb002"].Move(bax2+margin/2,gbay+Margin*2)
		noteGui["sb003"].GetPos(&bax3, &bay, &baw, &bah)
		noteGui["sb003"].Move(bax3+margin/2-2,gbay+Margin*2)
		
		noteGui["sb007"].GetPos(&bax, &bay, &baw, &bah)
		noteGui["sb007"].Move(bax+margin/2+1,Round(gbay+btnh+margin*2.3))
		noteGui["sb008"].GetPos(&bax, &bay, &baw, &bah)
		noteGui["sb008"].Move(bax+margin/2,Round(gbay+btnh+margin*2.3))
		noteGui["sb009"].GetPos(&bax, &bay, &baw, &bah)
		noteGui["sb009"].Move(bax+margin/2-2,Round(gbay+btnh+margin*2.3))
		
		noteGui.Add("GroupBox", "x" titleX+nttw+ntew+Margin " y" titleY+(btnh*3-Margin) " w" btnw*3+(Margin/2)*4 " h" btnh*6 " vgb2", "Action")
		noteGui["gb2"].GetPos(&gbax, &gbay, &gbaw, &gbah)	
		
		noteGui["sb013"].GetPos(&bax, &bay, &baw, &bah) ;search phone number
		noteGui["sb013"].Move(bax+margin/2+1,gbay+margin*2)
		noteGui["sb014"].GetPos(&bax, &bay, &baw, &bah) ;mcaid
		noteGui["sb014"].Move(bax+margin/2,gbay+margin*2)
		noteGui["sb015"].GetPos(&bax, &bay, &baw, &bah) ;hld30
		noteGui["sb015"].Move(bax+margin/2-2,gbay+margin*2)		
		
		noteGui["sb004"].GetPos(&bax, &bay, &baw, &bah)	;piflr
		noteGui["sb004"].Move(bax+margin/2+1,Round(gbay+btnh+margin*2.3))
		noteGui["sb005"].GetPos(&bax, &bay, &baw, &bah) ;slepa
		noteGui["sb005"].Move(bax+margin/2,Round(gbay+btnh+margin*2.3))
		noteGui["sb006"].GetPos(&bax, &bay, &baw, &bah) ;dispute
		noteGui["sb006"].Move(bax+margin/2-2,Round(gbay+btnh+margin*2.3))
		
		row2 := Round((gbay+btnh+margin*2.3)-(gbay+margin*2))
		row3 := Round((gbay+btnh+margin*2.3)+row2)
		
		noteGui["sb010"].GetPos(&bax, &bay, &baw, &bah) ;na
		noteGui["sb010"].Move(bax+margin/2+1,row3)
		noteGui["sb011"].GetPos(&bax, &bay, &baw, &bah) ;puhu
		noteGui["sb011"].Move(bax+margin/2,row3)
		noteGui["sb012"].GetPos(&bax, &bay, &baw, &bah) ;rtv
		noteGui["sb012"].Move(bax+margin/2-2,row3)
		
		row4 := row3+row2
		
		noteGui["sb016"].GetPos(&bax, &bay, &baw, &bah) ;NML
		noteGui["sb016"].Move(bax+margin/2+1,row4)
		noteGui["sb017"].GetPos(&bax, &bay, &baw, &bah) ;WN
		noteGui["sb017"].Move(bax+margin/2,row4)
		noteGui["sb018"].GetPos(&bax, &bay, &baw, &bah) ;TRC
		noteGui["sb018"].Move(bax+margin/2-2,row4)		
		
		row5 := row4+row2
		noteGui["sb019"].GetPos(&bax, &bay, &baw, &bah) ;AML
		noteGui["sb019"].Move(bax+margin/2+1,row5)
		noteGui["sb020"].GetPos(&bax, &bay, &baw, &bah) ;WORKING PQ
		noteGui["sb020"].Move(bax+margin/2,row5)
		noteGui["sb021"].GetPos(&bax, &bay, &baw, &bah) ;WORKING PQ
		noteGui["sb021"].Move(bax+margin/2-2,row5)
		;noteGui["gb2"].Move(,,,btnh*6)
		
		;MsgBox "row1 " gbay+margin*2 "`nrow2 " Round((gbay+btnh+margin*2.3)-(gbay+margin*2)) "`nor " Round((gbay+btnh+margin*2.3)) "`nor " row3 "`nor " Round((gbay+btnh+margin*2.3)-(gbay+margin*2))
		;noteGui["sb015"].GetPos(&bax, &bay, &baw, &bah)
		;noteGui["sb021"].Move(bax+margin/2,bay-btnh-margin/2+129)
		
		
		/****************
		* Small Buttons *
		****************/
		
;~~~~~~~name
		noteGui["e1"].GetPos(&tempx, &tempy, &tempw, &temph)
		noteGui.Add("Button","x" tempx " y" tempy " w" nteh " h" nteh " vimpb1 0x8000", ">").OnEvent("Click", imp1)
		noteGui["e1"].move(tempx+nteh+2,,tempw-nteh-2)
		imp1v := ""
		imp1(Button, Info){
		if WinExist("Latitude - Account Work Form"){
		try{
		LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form ahk_exe Latitude.exe")
		imp1v := LatitudeEl.ElementFromPath("YR4/").value
		}
		;MsgBox "1 " imp1v
		if imp1v = ""
		MsgBox "Could not copy name","Error",48+262144
		Else
		if InStr(imp1v,",") > 0{
		parts := StrSplit(imp1v, [", ", ","])
		lastName := parts[1]
		firstName := parts[2]
		;msgbox firstName " " lastName
		noteGui["e1"].value := firstName " " lastName
		}
		Else noteGui["e1"].value := imp1v
		 }
		Else
		MsgBox "Latitude not found","Error",48+262144
		}
;~~~~~~~phone
		noteGui["e3"].GetPos(&tempx, &tempy, &tempw, &temph)
		noteGui.Add("Button","x" tempx " y" tempy-1 " w" nteh " h" nteh+1 " vimpb2 0x8000", ">").OnEvent("Click", imp2)
		noteGui["e3"].move(tempx+nteh+2,,tempw-nteh-2)
		imp2(Button, Info){
		
		agntph := ""
		if WinExist("ahk_exe Agent Desktop Native.exe"){
		try{
		;MsgBox "try1 phone"
		AgentEl := UIA.ElementFromHandle("ahk_exe Agent Desktop Native.exe")
		if AgentEl.ElementFromPathExist("VKr"){
		agntph := AgentEl.ElementFromPath("VKr").Name
		agntph := RegExReplace(agntph, "\D", "")
		 }
		}
		
		if agntph = ""{
		try{
		AgentEl := UIA.ElementFromHandle("ahk_exe Agent Desktop Native.exe")
		if AgentEl.ElementFromPathExist("VKs"){
		agntph := AgentEl.ElementFromPath("VKs").Name
		agntph := RegExReplace(agntph, "\D", "")
		  }
		 }
		}
		
		if agntph = ""{		
		try{
		;MsgBox "try1 phone"
		AgentEl := UIA.ElementFromHandle("Agent Panel - Main window")
		if AgentEl.ElementFromPathExist("VKr"){
		agntph := AgentEl.ElementFromPath("VKr").Name
		agntph := RegExReplace(agntph, "\D", "")
		  }
		 }
		}
		
		if agntph = ""{
		try{
		AgentEl := UIA.ElementFromHandle("Agent Panel - Main window")
		if AgentEl.ElementFromPathExist("VKs"){
		agntph := AgentEl.ElementFromPath("VKs").Name
		agntph := RegExReplace(agntph, "\D", "")
		  }
		 }
		}
		
		;MsgBox agntph
		if agntph != ""{
		agntph := RegExReplace(agntph, "\D", "")
		noteGui["e3"].Value := agntph
		  }
		 Else
		 MsgBox "Could not get number from LiveVox","Error",48+262144
		 }
		Else
		MsgBox "Could not find LiveVox","Error",48+262144
		}
		
;~~~~~~~Email		
		noteGui["e4"].GetPos(&tempx, &tempy, &tempw, &temph)
		noteGui["e4"].move(tempx+nteh+2,,tempw-nteh*2-4)
		noteGui.Add("Button","x" tempx " y" tempy-1 " w" nteh " h" nteh+1 " vimpb3 0x8000", ">").OnEvent("Click", imp3)
		noteGui.Add("Button","x" tempx+tempw-nteh " y" tempy-1 " w" nteh " h" nteh+1 " vexpb3 0x8000", ">").OnEvent("Click", exp3)
		
		imp3(Button, Info){
		try{
		If WinExist("Latitude - Account Work Form"){
		goodbyepops()
		WinActivate "Latitude - Account Work Form"
		SendInput "!im"
		LatitudeE4 := UIA.ElementFromHandle("Latitude - Account Work Form")
		if LatitudeE4.ElementFromPath("RYYIYbQ/U").value = "(NULL)"
		MsgBox "No email","Error",48+262144
		else
		noteGui["e4"].Value := LatitudeE4.ElementFromPath("RYYIYbQ/U").value
		}
		else	
		MsgBox "Latitude not found","Error",48+262144
		}
		catch as e{
		MsgBox "An error has occured.","Error 0x" e.line,16+262144
		}
		}
		
		exp3(Button, Info){
		try{
		If WinExist("Latitude - Account Work Form"){
		if noteGui["e4"].Value = ""{
		MsgBox "Can not export blank email. Please fill in email and try again.","Error be0x878",16+262144
		Return
		}
		if !InStr(noteGui["e4"].Value,"@"){
		MsgBox "Please check the email and try again.","Error be0x874",16+262144
		Return
		}
		goodbyepops()
		WinActivate "Latitude - Account Work Form"
		SendInput "!im"
		LatitudeE4 := UIA.ElementFromHandle("Latitude - Account Work Form")
		;emailimp := 
		;msgbox noteGui["e4"].Value
		LatitudeE4.ElementFromPath("RYYIYbQ/U").value := noteGui["e4"].Value
		LatitudeE4.ElementFromPath("RYYIY0").ControlClick()
		sleep 100
		SendInput "!y"
		sleep 200
		WinClose ("Successful ahk_exe Latitude.exe")
		;msgbox noteGui["e4"].Value "posted"
		}
		else	
		MsgBox "Latitude not found","Error",48+262144
		}
		catch as e{
		MsgBox "An error has occured.","Error 0x" e.line,16+262144
		}
		}
		
;~~~~~~~adv bal
		noteGui["e5"].GetPos(&tempx, &tempy, &tempw, &temph)
		noteGui.Add("Button","x" tempx " y" tempy-1 " w" nteh " h" nteh+1 " vimpb4 0x8000", ">").OnEvent("Click", imp4)
		noteGui["e5"].move(tempx+nteh+2,,tempw-nteh-2)
		imp4(Button, Info){
		goodbyepops()
		CoordMode "Mouse", "Client"
		; MouseGetPos &balmoux, &balmouy
		; msgbox balmoux " & " balmouy
		if WinExist("Account Balance Information"){
		WinShow "Account Balance Information"
		WinActivate "Account Balance Information"
		LatitudeE2 := UIA.ElementFromHandle("Account Balance Information")
		balstr := LatitudeE2.ElementFromPath("R4s").value
		balstr := RegExReplace(balstr,"[\$,]", "")
		noteGui["e5"].value := balstr
		}
		else{
		try{
		CoordMode "Mouse", "Screen"
		MouseGetPos &mousex, &mousey
		;msgbox "x " mousex " y " mousey
		noteGui["impb4"].GetPos(&tempbx, &tempby, &tempbw, &tempbh)
		;msgbox "x " tempbx " & y " tempby
		if WinExist("Latitude - Account Work Form"){
		goodbyepops()
		WinActivate "Latitude - Account Work Form"
		CoordMode "Mouse", "Client"
		MouseClick , 975*WinGetDpi("Latitude - Account Work Form")/144, 466*WinGetDpi("Latitude - Account Work Form")/144
		DllCall("SetCursorPos", "int", mousex, "int", mousey)
		sleep 100
		LatitudeE2 := UIA.ElementFromHandle("Account Balance Information")
		balstr := LatitudeE2.ElementFromPath("R4s").value
		balstr := RegExReplace(balstr,"[\$,]", "")
		noteGui["e5"].value := balstr
		WinClose("Account Balance Information ahk_exe Latitude.exe")
		}
		Else
		MsgBox "Latitude not found","Error",48+262144
		}
		Catch{
		MsgBox "Unable to access balance. Please try again.","Error",16+262144
  }
 }
}
		
/************
*    EPA    *
************/

		noteGui["Title"].GetPos(&titleX, &titleY, &titleW, &titleH)		
		noteGui.Add("Text","x" titleX+nttw+ntew+Margin " y" Margin " w" nttw+ntew+margin*2 " h" Margin*3 " 0x200 Center Border vEPATitle","EPA")
		noteGui["EPATitle"].SetFont("S" Margin*1.75 " w" Margin*100)
		
		For con_name, header in epacMap{	
				epaanch := A_Index*margin*2.97
				noteGui.Add("Text","x" titleX+nttw+ntew+Margin " y" titleY+margin+epaanch " w" nttw+50 " h" ntth " 0x200 vepat" con_name, header)
				noteGui.Add("Edit","xp+" nttw+margin*2 " w" ntew " h" nteh " Uppercase 0x200 vepae" con_name,"")
		} ;End Loop
		
		
		/************
		*	OK CB	*
		************/
		noteGui["epae9"].Visible := 0
		noteGui.Add("Text","xp yp+" Margin/5 " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcbOK1 Center BackgroundWhite border","")
		noteGui.Add("Text","xp+" nteh*0.90+Margin/2 " h" nteh*0.75 " 0x200 vtY","Yes")
		noteGui.Add("Text","xp+" nttw " yp w" nteh*0.75 " h" nteh*0.75 " 0x200 vcbOK2 Center BackgroundWhite border","")
		noteGui.Add("Text","xp+" nteh*0.90+Margin/2 " h" nteh*0.75 " 0x200 vtN","No")
		
		noteGui["tY"].OnEvent("Click", cbYCB)
		noteGui["cbOK1"].OnEvent("Click", cbYCB)
		cbYCB(Text, Info){
			If noteGui["cbOK1"].Text = checkmark
			noteGui["cbOK1"].Text := ""
			Else If noteGui["cbOK1"].Text = ""
			noteGui["cbOK1"].Text := checkmark
			noteGui["cbOK2"].Text := ""
		} ;End cbNCB

		noteGui["tN"].OnEvent("Click", cbNCB)
		noteGui["cbOK2"].OnEvent("Click", cbNCB)
		cbNCB(Text, Info){
			If noteGui["cbOK2"].Text = checkmark
			noteGui["cbOK2"].Text := ""
			Else If noteGui["cbOK2"].Text = ""
			noteGui["cbOK2"].Text := checkmark
			noteGui["cbOK1"].Text := ""
		} ;End cbNCB
		
		/************
		*	RECUR	*
		************/
		noteGui["epae7"].Visible := 0
		noteGui["epae7"].GetPos(&X6, &Y6, &W6, &H6)
		noteGui.Add("Text","x" X6 " y" Y6+3 " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcbRec Center BackgroundWhite border","")
		noteGui["cbRec"].OnEvent("Click", cbrec)
		noteGui.Add("DropDownList","xp+" Margin+nteh*0.75 " yp-" Margin/5 " w" ntew-nteh*0.75-margin " h" nteh " Disabled Uppercase 0x200 r10 vepadd",["","WEEKLY","EVERY TWO WEEKS","TWICE PER MONTH","EVERY 28 DAYS","MONTHLY","END OF EA. MONTH","LAST FRI. EA. MONTH","EVERY OTHER MONTH","BI-WEEKLY"])
		
		cbrec(Text, Info){
			If noteGui["cbRec"].Text = checkmark{
				noteGui["cbRec"].Text := ""
				noteGui["epadd"].Enabled := 0
				noteGui["epadd"].Text := ""
				}
			Else If noteGui["cbRec"].Text = ""{
				noteGui["cbRec"].Text := checkmark
				noteGui["epadd"].Enabled := 1
				}
			}
			
		
		/******************
		*	EPA BUTTONS   *
		******************/
		noteGui["b9"].GetPos(&rx,&ry,&rw,&b9h)
		noteGui.Add("Button","x" rx+btnw+margin*1.5 " y" ry " w" btnw " h" btnh " vepab1", "Close EPA").OnEvent("Click", CloseEPAb)
		noteGui.Add("Button","x" rx+btnw*2+margin*2.5 " y" ry " w" btnw " h" btnh " vepab2", "Disclosure").OnEvent("Click", DiclosureEPA)
		noteGui.Add("Button","x" rx+btnw*3+margin*3.5 " y" ry " w" btnw " h" btnh " vepab3", "Post EPA").OnEvent("Click", PostEPA)
				
		/****************
		* Small Buttons *
		****************/
		
		noteGui["epae1"].GetPos(&epat1x,&epat1y,&epat1w,&epat1h) ;name
		noteGui.Add("Button","x" epat1x " y" epat1y " w" nteh " h" nteh " vepab4 0x8000", ">").OnEvent("Click", EPAimp1)
		EPAimp1(Button, Info){
		noteGui["epae1"].value := noteGui["e1"].value
		}
		
		noteGui["epae4"].GetPos(&epat5x,&epat5y,&epat5w,&epat5h) ;bank name
		noteGui.Add("Button","x" epat5x " y" epat5y+1 " w" nteh " h" nteh " vepab5", ">").OnEvent("Click", EPAimp2)
		
		noteGui["epae5"].GetPos(&epat5x,&epat5y,&epat5w,&epat5h) ;ammt
		noteGui.Add("Button","x" epat5x " y" epat5y+1 " w" nteh " h" nteh " vepab6", ">").OnEvent("Click", EPAimp3)
		EPAimp3(Button, Info){
		noteGui["epae6"].value := noteGui["e5"].value
		}

/***************************************
* Each chunk is defines our checkboxes *
* in Main and our functions for them   *
***************************************/

			noteGui["tDOB"].OnEvent("Click", cbMMc)
			noteGui["cb1DOB"].OnEvent("Click", cbMMc)
			cbMMc(Text, Info){
			If noteGui["cb1DOB"].Text = checkmark
			noteGui["cb1DOB"].Text := ""
			,cbMMv := ""
			Else If noteGui["cb1DOB"].Text = ""
			noteGui["cb1DOB"].Text := checkmark
			} ;End cbMMc
					
			noteGui["tSSN"].OnEvent("Click", cbDOBc)
			noteGui["cb2SSN"].OnEvent("Click", cbDOBc)
			cbDOBc(Text, Info){
			If noteGui["cb2SSN"].Text = checkmark
			noteGui["cb2SSN"].Text := ""
			,cbDOBv := ""
			Else If noteGui["cb2SSN"].Text = ""
			noteGui["cb2SSN"].Text := checkmark
			} ;End cbDOBc
					
			noteGui["tMM"].OnEvent("Click", cbSSNc)
			noteGui["cb3MM"].OnEvent("Click", cbSSNc)
			cbSSNc(Text, Info){
			If noteGui["cb3MM"].Text = checkmark
			noteGui["cb3MM"].Text := ""
			,cbSSNv := ""
			Else If noteGui["cb3MM"].Text = ""
			noteGui["cb3MM"].Text := checkmark
			} ;End cbSSNc
						
			noteGui["tdisp"].OnEvent("Click", cbdispc)
			noteGui["cbdisp"].OnEvent("Click", cbdispc)
			cbdispc(Text, Info){
			If noteGui["cbdisp"].Text = checkmark
			noteGui["cbdisp"].Text := ""
			Else If noteGui["cbdisp"].Text = ""
			noteGui["cbdisp"].Text := checkmark
			} ;End cbdispc
			
			noteGui["tcease"].OnEvent("Click", cbceasec)
			noteGui["cbcease"].OnEvent("Click", cbceasec)
			cbceasec(Text, Info){
			If noteGui["cbcease"].Text = checkmark
			noteGui["cbcease"].Text := ""
			Else If noteGui["cbcease"].Text = ""
			noteGui["cbcease"].Text := checkmark
			} ;End cbdispc			
			
			noteGui["tbad"].OnEvent("Click", cbbadc)
			noteGui["cbbad"].OnEvent("Click", cbbadc)
			cbbadc(Text, Info){
			If noteGui["cbbad"].Text = checkmark
			noteGui["cbbad"].Text := ""
			Else If noteGui["cbbad"].Text = ""
			noteGui["cbbad"].Text := checkmark
			noteGui["cbconsent"].Text := ""
			} ;End cbbadc			
			
			noteGui["tconsent"].OnEvent("Click", cbconsentc)
			noteGui["cbconsent"].OnEvent("Click", cbconsentc)
			cbconsentc(Text, Info){
			If noteGui["cbconsent"].Text = checkmark
			noteGui["cbconsent"].Text := ""
			Else If noteGui["cbconsent"].Text = ""
			noteGui["cbconsent"].Text := checkmark
			noteGui["cbbad"].Text := ""
			} ;End cbconsentc			
			
		if settings["esls"] = 1{			
			noteGui["cb4ls1"].OnEvent("Click", cbls1c)
			noteGui["tls1"].OnEvent("Click", cbls1c)
			cbls1c(Text, Info){
			If noteGui["cb4ls1"].Text = checkmark
			noteGui["cb4ls1"].Text := ""
			Else If noteGui["cb4ls1"].Text = ""
			noteGui["cb4ls1"].Text := checkmark
			} ;End cbls1c
		}
		
		if settings["esls"] = 1{					
			noteGui["cb5ls2"].OnEvent("Click", cbls2c)
			noteGui["tls2"].OnEvent("Click", cbls2c)
			cbls2c(Text, Info){
			If noteGui["cb5ls2"].Text = checkmark
			noteGui["cb5ls2"].Text := ""
			Else If noteGui["cb5ls2"].Text = ""
			noteGui["cb5ls2"].Text := checkmark
			} ;End cbls2c			
		}
			
			
/***************************************
* This is what updates the other boxes *
*    when typing an account number     *
***************************************/

noteGui["epae2"].OnEvent("Change", BINupd) 	;updates some boxes


		BINupd(Edit, Info){
		
		aclStr := StrLen(noteGui["epae2"].Value)
		if aclStr<16{
			noteGui["epae3"].Text := "N/A"
			}
		
		if aclStr=16{
			 BINa := SubStr(noteGui["epae2"].Text,1,6)
				noteGui["epae3"].Text := BINa
			}
		
		if aclStr>16{
			noteGui["epae3"].Text := "Error"
			}
		}

noteGui["epae2"].OnEvent("Change", l4upd)	;updates some boxes

		l4upd(Edit, Info){
			L4a := SubStr(noteGui["epae2"].Text,-4,4)
			noteGui["epae4"].Text := L4a		
		
		}

/**************************
*	Need to declare some  * 
*	variables here		  *
**************************/

noteGui["epat9"].GetPos(&epat9x,&epat9y,&epat9w,&epat9h)
noteGui["cbOK1"].GetPos(&cbOK1x,&cbOK1y,&cbOK1w,&cbOK1h)
noteGui["cbOK2"].GetPos(&cbOK2x,&cbOK2y,&cbOK2w,&cbOK2h)
noteGui["tY"].GetPos(&tYx,&tYy,&tYw,&tYh)
noteGui["tN"].GetPos(&tNx,&tNy,&tNw,&tNh)
Global cbOK1x, cbOK2x, tYx, tNx, cbOK1y, cbOK2y, tYy, tNy, epat9x, epat9y


/************************************
*   Moving some stuff around here   *
************************************/
noteGui["epae2"].Opt("Number")
noteGui["epae3"].Opt("ReadOnly")
noteGui["epae4"].Opt("ReadOnly")

noteGui["epae1"].GetPos(&epae1x,&epae1y,&epae1w,&epae1h)
noteGui["epae1"].Move(epae1x+nteh+2,,epae1w-nteh-2)

If settings["esls"] = 0{

noteGui["epat3"].GetPos(&epat3x,&epat3y,&epat3w,&epat3h)
noteGui["epat3"].Move(,epat3y-Margin/8,,)
noteGui["epae3"].Move(epat3x+Margin*3,epat3y-Margin/8,Margin*10)
noteGui["epat4"].Move(epat3x+Margin*15,epat3y-Margin/8,Margin*10)
noteGui["epae4"].Move(epat3x+Margin*20,epat3y-Margin/8,Margin*10)

noteGui["e3"].GetPos(&tempx,&tempy,&tempw,&temph) ;bank name
noteGui["epat5"].Move(tempx+tempw+Margin,tempy)
noteGui["epae5"].Move(tempx+tempw+nttw+Margin*3,tempy,ntew-nteh-2)
noteGui["epae5"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epab5"].Move(tempx+tempw+2,tempy)

noteGui["e4"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epat6"].Move(tempx+tempw+Margin+nteh,tempy)
noteGui["epae6"].Move(tempx+tempw+nttw+Margin*3+nteh*2+4,tempy,tempw+nteh)
noteGui["epae6"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epab6"].Move(,tempy)

noteGui["e5"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epat7"].Move(tempx+tempw+Margin,tempy)
noteGui["cbRec"].Move(tempx+tempw+nttw+Margin*3,tempy+4)
noteGui["epadd"].Move(tempx+tempw+nttw+nteh*0.75+Margin*4,tempy+1)

noteGui["e6"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epat8"].Move(tempx+tempw+Margin,tempy)
noteGui["epae8"].Move(tempx+tempw+nttw+Margin*3,tempy)

noteGui["t7"].GetPos(&temptx,&tempty,&temptw,&tempth)
noteGui["cbdisp"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epat9"].Move(temptx+temptw+Margin,tempy)
noteGui["cbOK1"].Move(temptx+temptw+nttw+Margin*5,tempy+Margin/4)
noteGui["tY"].Move(temptx+temptw+nttw+Margin*7,tempy+Margin/4)
noteGui["cbOK2"].Move(temptx+temptw+nttw+Margin*17,tempy+Margin/4)
noteGui["tN"].Move(temptx+temptw+nttw+Margin*19,tempy+Margin/4)

}

Else if settings["esls"] = 1{

noteGui["epat3"].GetPos(&epat3x,&epat3y,&epat3w,&epat3h)
noteGui["epat3"].Move(,epat3y-Margin/8,,)
noteGui["epae3"].Move(epat3x+Margin*3,epat3y-Margin/8,Margin*10)
noteGui["epat4"].Move(epat3x+Margin*15,epat3y-Margin/8,Margin*10)
noteGui["epae4"].Move(epat3x+Margin*20,epat3y-Margin/8,Margin*10)

noteGui["tls2"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epat5"].Move(tempx+tempw+Margin,tempy)
noteGui["epae5"].Move(tempx+tempw+nttw+Margin*3,tempy,)

noteGui["epae5"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epae5"].Move(,,tempw-nteh-2)
noteGui["epab5"].Move(tempx+tempw-nteh,tempy)

noteGui["els3"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epat6"].Move(tempx+tempw+Margin,tempy)
noteGui["epae6"].Move(tempx+tempw+nttw+nteh+Margin*3,tempy)
noteGui["epab6"].Move(tempx+tempw+nttw+Margin*3-2,tempy)

noteGui["epae5"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epae6"].Move(,,tempw)

noteGui["e3"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epat7"].Move(tempx+tempw+Margin,tempy)
noteGui["cbRec"].Move(tempx+tempw+nttw+Margin*3,tempy+4)
noteGui["epadd"].Move(tempx+tempw+nttw+nteh*0.75+Margin*4,tempy+1)

noteGui["e4"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epat8"].Move(tempx+tempw+Margin+nteh,tempy)
noteGui["epae8"].Move(tempx+tempw+nttw+Margin*3+nteh,tempy)

noteGui["e5"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epat9"].Move(tempx+tempw+Margin,tempy)
noteGui["cbOK1"].Move(tempx+tempw+nttw+Margin*5,tempy+Margin/4)
noteGui["tY"].Move(tempx+tempw+nttw+Margin*7,tempy+Margin/4)
noteGui["cbOK2"].Move(tempx+tempw+nttw+Margin*17,tempy+Margin/4)
noteGui["tN"].Move(tempx+tempw+nttw+Margin*19,tempy+Margin/4)

}


/************
*    F/C    *
************/

	; Global fccbMap := Map(
	; "cb1ms",	"Single"		,
	; "cb2ms",	"Married"		,
	; "cb3ro",	"Rent"			,
	; "cb4ro",	"Own"			,
	; "cb5em",	"Unemployed"	,
	; "cb6em",	"Employed"		,
	; "cb7bn",	"Refused"		,
	; "cb8bn",	"Consent"		, 
	; )
	
			; noteGui["tDOB"].OnEvent("Click", cbMMc)
			; noteGui["cb1DOB"].OnEvent("Click", cbMMc)
			; cbMMc(Text, Info){
			; If noteGui["cb1DOB"].Text = checkmark
			; noteGui["cb1DOB"].Text := ""
			; ,cbMMv := ""
			; Else If noteGui["cb1DOB"].Text = ""
			; noteGui["cb1DOB"].Text := checkmark
			; } ;End cbMMc
	
	; If A_Index = 2{ ;Makes the CheckBoxes
	; noteGui.Add("Text","xm+" nttw " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb1DOB Center BackgroundWhite border","")
	; noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtDOB","DOB")
	; noteGui.Add("Text","xp+" nttw/1.3 " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb2SSN Center BackgroundWhite border","")
	; noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtSSN","SSN")
	; noteGui.Add("Text","xp+" nttw/1.255 " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb3MM Center BackgroundWhite border","")
	; noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtMM","MM")

		noteGui["Title"].GetPos(&titleX, &titleY, &titleW, &titleH)		
		noteGui.Add("Text","x" titleX+nttw+ntew+Margin " y" Margin " w" nttw+ntew+margin*3 " h" Margin*3 " 0x200 Center Border vFCTitle","Full and Complete")
		noteGui["FCTitle"].SetFont("S" Margin*1.75 " w" Margin*100)
		
		For con_name, header in fccMap{	
				epaanch := A_Index*margin*2.97
				noteGui.Add("Text","x" titleX+nttw+ntew+Margin " y" titleY+margin+epaanch+(nteh*A_index) " w" nttw+50 " h" ntth " 0x200 vfct" con_name, header)
				;noteGui.Add("Edit","xp+" nttw+margin*3 " w" ntew " h" nteh " Uppercase 0x200 vfce" con_name,"")
		} ;End Loop

;buttons		
		noteGui["b9"].GetPos(&rx,&ry,&rw,&b9h)
		noteGui.Add("Button","x" rx+btnw+margin*1.5 " y" ry " w" btnw " h" btnh " vfcb1", "Close F/C").OnEvent("Click", CloseFCb)
		noteGui.Add("Button","x" rx+btnw*1.82+margin*5 " y" ry " w" btnw " h" btnh " vfcb2", "Fill F/C").OnEvent("Click", FillFC)
		noteGui.Add("Button","x" rx+btnw*3+margin*5 " y" ry " w" btnw " h" btnh " vfcb3", "Post F/C").OnEvent("Click", PostFC)

;checkbox
	;row 1	
		noteGui["fct1"].GetPos(&tempx,&tempy,&tempw,&temph)
		
		noteGui.Add("Text","x" tempx+nttw+margin*3 " y" tempy+2 " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb1ms Center BackgroundWhite border","")
		noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtcb1ms","Single")
		noteGui["cb1ms"].OnEvent("Click", cb1MSc)
		noteGui["tcb1ms"].OnEvent("Click", cb1MSc)
		cb1MSc(Text, Info){
		If noteGui["cb1ms"].Text = checkmark
		noteGui["cb1ms"].Text := ""
		Else If noteGui["cb1ms"].Text = ""
		noteGui["cb1ms"].Text := checkmark
		noteGui["cb2ms"].Text := ""
		noteGui["fceMS"].Enabled := 0
		} ;End cb1MSc
		
		noteGui.Add("Text","xp+" nttw " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb2ms Center BackgroundWhite border","")
		noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtcb2ms","Married")
		noteGui.Add("Edit","x" tempx+nttw+27 " y" tempy+ntth+4 " w" titleW-nttw+72  " h" nteh " Disabled Uppercase 0x200 vfceMS","")
		noteGui["cb2ms"].OnEvent("Click", cb2MSc)
		noteGui["tcb2ms"].OnEvent("Click", cb2MSc)
		
		cb2MSc(Text, Info){
		If noteGui["cb2ms"].Text = checkmark{
		noteGui["cb2ms"].Text := ""
		noteGui["fceMS"].Enabled := 0
		}
		Else If noteGui["cb2ms"].Text = ""{
		noteGui["cb2ms"].Text := checkmark
		noteGui["cb1ms"].Text := ""
		noteGui["fceMS"].Enabled := 1
		}
		} ;End cb2MSc

noteGui["fct1"].Move(,tempy-20)
noteGui["cb1ms"].Move(,tempy-20)
noteGui["tcb1ms"].Move(,tempy-20)
noteGui["cb2ms"].Move(,tempy-20)
noteGui["tcb2ms"].Move(,tempy-20)
noteGui["fceMS"].Move(,tempy)
	
	
	;row 2	
		noteGui["fct2"].GetPos(&tempx,&tempy,&tempw,&temph)
		
		noteGui.Add("Text","x" tempx+nttw+margin*3 " y" tempy+2 " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb3ro Center BackgroundWhite border","")
		noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtcb3ro","Rent")
		noteGui["cb3ro"].OnEvent("Click", cb3roc)
		noteGui["tcb3ro"].OnEvent("Click", cb3roc)
		cb3roc(Text, Info){
		If noteGui["cb3ro"].Text = checkmark
		noteGui["cb3ro"].Text := ""
		Else If noteGui["cb3ro"].Text = ""
		noteGui["cb3ro"].Text := checkmark
		noteGui["cb4ro"].Text := ""
		; noteGui["fceRO"].Enabled := 0
		} ;End cb3roc
		
		noteGui.Add("Text","xp+" nttw " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb4ro Center BackgroundWhite border","")
		noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtcb4ro","Own")
		; noteGui.Add("Edit","x" tempx+nttw+27 " y" tempy+ntth+4 " w" titleW-nttw+72  " h" nteh " Disabled Uppercase 0x200 vfceRO","")
		noteGui["cb4ro"].OnEvent("Click", cb4roc)
		noteGui["tcb4ro"].OnEvent("Click", cb4roc)
		cb4roc(Text, Info){
		If noteGui["cb4ro"].Text = checkmark{
		noteGui["cb4ro"].Text := ""
		; noteGui["fceRO"].Enabled := 0
		}
		Else If noteGui["cb4ro"].Text = ""{
		noteGui["cb4ro"].Text := checkmark
		noteGui["cb3ro"].Text := ""
		; noteGui["fceRO"].Enabled := 1
		}
		} ;End cb4roc

noteGui["fct2"].Move(,tempy-20)
noteGui["cb3ro"].Move(,tempy-20)
noteGui["tcb3ro"].Move(,tempy-20)
noteGui["cb4ro"].Move(,tempy-20)
noteGui["tcb4ro"].Move(,tempy-20)
; noteGui["fceRO"].Move(,tempy)

	;row 3	
		noteGui["fct3"].GetPos(&tempx,&tempy,&tempw,&temph)	

		noteGui.Add("Text","x" tempx+nttw+margin*3 " y" tempy+2 " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb5em Center BackgroundWhite border","")
		noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtcb5em","Unemployed")
		
		noteGui["cb5em"].OnEvent("Click", cb5emc)
		noteGui["tcb5em"].OnEvent("Click", cb5emc)
		cb5emc(Text, Info){
		If noteGui["cb5em"].Text = checkmark
		noteGui["cb5em"].Text := ""
		Else If noteGui["cb5em"].Text = ""
		noteGui["cb5em"].Text := checkmark
		noteGui["cb6em"].Text := ""
		noteGui["fceEM"].Enabled := 0
		} ;End cb5emc

		noteGui.Add("Text","xp+" nttw " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb6em Center BackgroundWhite border","")
		noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtcb6em","Employed")
		noteGui.Add("Edit","x" tempx+nttw+27 " y" tempy+ntth+4 " w" titleW-nttw+72  " h" nteh " Disabled Uppercase 0x200 vfceEM","")
		noteGui["cb6em"].OnEvent("Click", cb6emc)
		noteGui["tcb6em"].OnEvent("Click", cb6emc)
		cb6emc(Text, Info){
		If noteGui["cb6em"].Text = checkmark{
		noteGui["cb6em"].Text := ""
		noteGui["fceEM"].Enabled := 0
		}
		Else If noteGui["cb6em"].Text = ""{
		noteGui["fceEM"].Enabled := 1
		noteGui["cb6em"].Text := checkmark
		noteGui["cb5em"].Text := ""
		}
		} ;End cb6emc

noteGui["fct3"].Move(,tempy-20-nteh)
noteGui["cb5em"].Move(,tempy-20-nteh)
noteGui["tcb5em"].Move(,tempy-20-nteh)
noteGui["cb6em"].Move(,tempy-20-nteh)
noteGui["tcb6em"].Move(,tempy-20-nteh)
noteGui["fceEM"].Move(,tempy-nteh)

	;row 4
		noteGui["fct4"].GetPos(&tempx,&tempy,&tempw,&temph)	

		noteGui.Add("Text","x" tempx+nttw+margin*3 " y" tempy+2 " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb7bn Center BackgroundWhite border","")
		noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtcb7bn","Refused")

		noteGui["cb7bn"].OnEvent("Click", cb7bnc)
		noteGui["tcb7bn"].OnEvent("Click", cb7bnc)
		cb7bnc(Text, Info){
		If noteGui["cb7bn"].Text = checkmark
		noteGui["cb7bn"].Text := ""
		Else If noteGui["cb7bn"].Text = ""
		noteGui["cb7bn"].Text := checkmark
		noteGui["cb8bn"].Text := ""
		noteGui["fceBN"].Enabled := 0
		} ;End cb7bnc

		noteGui.Add("Text","xp+" nttw " w" nteh*0.75 " h" nteh*0.75 " 0x200 vcb8bn Center BackgroundWhite border","")
		noteGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtcb8bn","Consent")
		noteGui.Add("Edit","x" tempx+nttw+27 " y" tempy+ntth+4 " w" titleW-nttw+72  " h" nteh " Disabled Uppercase 0x200 vfceBN","")
		noteGui["cb8bn"].OnEvent("Click", cb8bnc)
		noteGui["tcb8bn"].OnEvent("Click", cb8bnc)
		cb8bnc(Text, Info){
		If noteGui["cb8bn"].Text = checkmark{
		noteGui["cb8bn"].Text := ""
		noteGui["fceBN"].Enabled := 0
		}
		Else If noteGui["cb8bn"].Text = ""{
		noteGui["fceBN"].Enabled := 1
		noteGui["cb8bn"].Text := checkmark
		noteGui["cb7bn"].Text := ""
		}
		} ;End cb7bnc

noteGui["fct4"].Move(,tempy-20-nteh)
noteGui["cb7bn"].Move(,tempy-20-nteh)
noteGui["tcb7bn"].Move(,tempy-20-nteh)
noteGui["cb8bn"].Move(,tempy-20-nteh)
noteGui["tcb8bn"].Move(,tempy-20-nteh)
noteGui["fceBN"].Move(,tempy-nteh)		

		
/**************
*    SLEPA    *
***************/

		noteGui["Title"].GetPos(&titleX, &titleY, &titleW, &titleH)		
		noteGui.Add("Text","x" titleX+nttw+ntew+Margin " y" Margin " w" nttw+ntew+margin*3 " h" Margin*3 " 0x200 Center Border vSLEPATitle","SLEPA")
		noteGui["SLEPATitle"].SetFont("S" Margin*1.75 " w" Margin*100)
		
		For con_name, header in slepaMap{	
				epaanch := A_Index*margin*2.97
				noteGui.Add("Text","x" titleX+nttw+ntew+Margin " y" titleY+margin+epaanch " w" nttw+50 " h" ntth " 0x200 vslepat" con_name, header)
				noteGui.Add("Edit","xp+" nttw+margin*3+10 " w" ntew-10 " h" nteh " Uppercase 0x200 vslepae" con_name,"")
		} ;End Loop
		
		noteGui["b9"].GetPos(&rx,&ry,&rw,&b9h)

		noteGui.Add("Button","x" rx+btnw+margin*1.5 " y" ry " w" btnw " h" btnh " vslepab1", "Close SLEPA").OnEvent("Click", CloseSLEPAb)

		noteGui["slepae1"].GetPos(&tempx,&tempy,&tempw,&temph)
		noteGui.Add("Button","x" tempx " y" tempy " w" nteh " h" nteh " vslepab2", ">").OnEvent("Click", SLEPAimp1)
		SLEPAimp1(Button, Info){
		noteGui["slepae1"].Value := noteGui["e1"].Value
		}
		noteGui["slepae1"].move(tempx+nteh+3,,tempw-nteh-3)

		noteGui.Add("Button","x" rx+btnw*3+margin*5 " y" ry " w" btnw " h" btnh " vslepab3", "Post SLEPA").OnEvent("Click", PostSLEPA)
		
		noteGui["slepae3"].GetPos(&tempx,&tempy,&tempw,&temph)
		noteGui.Add("Button","x" tempx " y" tempy " w" nteh " h" nteh " vslepab4", ">").OnEvent("Click", SLEPAimp2)
		SLEPAimp2(Button, Info){
		noteGui["slepae3"].Value := noteGui["e4"].Value
		}
		noteGui["slepae3"].move(tempx+nteh+3,,tempw-nteh-3)
		
/*
**
*/

noteGui["Title"].GetPos(&titleX, &titleY, &titleW, &titleH)		
noteGui.Add("Text","x" titleX+nttw+ntew+Margin " y" Margin " w" nttw+ntew+margin " h" Margin*3 " 0x200 Center Border vatytitle","ATTORNEY NOTE")
noteGui["atytitle"].SetFont("S" margin*1.6 " w" Margin*100)

noteGui.Add("Text","x" titleX+nttw+ntew+Margin " y" titleY+titleH+Margin " w" nttw " h" ntth " 0x200 vatyt1", "ATY TYPE:")
noteGui.Add("Edit","yp w" ntew " h" nteh " Uppercase 0x200 vatye1","")
noteGui.Add("Text","x" titleX+nttw+ntew+Margin " y" titleY+titleH+Margin*2+ntth " w" nttw " h" ntth " 0x200 vatyt2", "ATY NAME:")
noteGui.Add("Edit","yp w" ntew " h" nteh " Uppercase 0x200 vatye2","")
noteGui.Add("Text","x" titleX+nttw+ntew+Margin " y" titleY+titleH+Margin*3+ntth*2 " w" nttw " h" ntth " 0x200 vatyt3", "ATY PH#:")
noteGui.Add("Edit","yp w" ntew " h" nteh " Uppercase 0x200 vatye3","")
noteGui.Add("Text","x" titleX+nttw+ntew+Margin " y" titleY+titleH+Margin*4+ntth*3 " w" nttw " h" ntth " 0x200 vatyt4", "ATY FAX#:")
noteGui.Add("Edit","yp w" ntew " h" nteh " Uppercase 0x200 vatye4","")

notegui["atye4"].GetPos(&atyx, &atyy, &atyw, &atyh)

;~~~~close aty
noteGui.Add("Button","x" rx+btnw+margin*1.5 " y" ry " w" btnw " h" btnh " vatyb1", "Close ATY").OnEvent("Click", CloseATYb)
CloseATYb(Button, Info){
Result := MsgBox("Would you like to close the Attorney Note?","Close ATY?",32+262144 " YesNo")
if Result = "No"
Return
Else
notegui["atytitle"].Visible := 0
loop 4{
notegui["atyt" A_index].Visible := 0
notegui["atye" A_index].Value   := ""
notegui["atye" A_index].Visible := 0
}
notegui["atyb1"].Visible := 0
notegui["atyb2"].Visible := 0
noteGui.Hide
noteGui.Show("Autosize")
}

CloseATY(){
notegui["atytitle"].Visible := 0
loop 4{
notegui["atyt" A_index].Visible := 0
notegui["atye" A_index].Value   := ""
notegui["atye" A_index].Visible := 0
}
notegui["atyb1"].Visible := 0
notegui["atyb2"].Visible := 0
}

;~~~~~~~~postaty
noteGui.Add("Button","x" atyx+atyw-btnw " y" ry " w" btnw " h" btnh " vatyb2", "Post ATY").OnEvent("Click", PostATY)
PostATY(Button, Info){
atystr := "ATTORNEY TYPE:,ATTORNEY NAME:,ATTORNEY PHONE NUMBER:,ATTORNEY FAX NUMBER:"
atyout := ""
loop parse, atystr, ","{
if notegui["atye" A_index].Value = ""
atyout .= A_LoopField ", "
Else
atyout .= A_LoopField " " notegui["atye" A_index].value ", "
}
Result := MsgBox("Would you like to post the Attorney Note?","Post ATY?",32+262144 " YesNo")
if Result = "No"
Return
Else
	If WinExist("Latitude - Account Work Form"){
		WinActivate "Latitude - Account Work Form"
		SendInput "!n"
		sleep 250
		SendInput "!n"
		WinWait "Latitude - New Note"
		If WinExist("Latitude - New Note"){
			WinActivate "Latitude - New Note"
			A_Clipboard := RTrim(atyout, ", ")
			WinActivate "Latitude - New Note"
			SendInput "CO{Down}{Tab}ATY{Down}{Tab}"
			SendInput "^+{End}" "{Del}"
			SendInput "^v"
			SendInput "!a"
			atyout := ""
		}
	}
	else	
	MsgBox "Latitude not found","Error",48+262144
}








/*********************
*    Phone Search    *
*********************/		

PhnSrch(){
if WinExist("Latitude - Account Work Form"){
goodbyepops()
agntph := ""
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
if noteGui["e3"].Value = ""{
if WinExist("ahk_exe Agent Desktop Native.exe"){
try{
;MsgBox "try1 phone"
AgentEl := UIA.ElementFromHandle("ahk_exe Agent Desktop Native.exe")
if AgentEl.ElementFromPathExist("VKr"){
agntph := AgentEl.ElementFromPath("VKr").Name
agntph := RegExReplace(agntph, "\D", "")
 }
}

if agntph = ""{
try{
AgentEl := UIA.ElementFromHandle("ahk_exe Agent Desktop Native.exe")
if AgentEl.ElementFromPathExist("VKs"){
agntph := AgentEl.ElementFromPath("VKs").Name
agntph := RegExReplace(agntph, "\D", "")
  }
 }
}

if agntph = ""{		
try{
;MsgBox "try1 phone"
AgentEl := UIA.ElementFromHandle("Agent Panel - Main window")
if AgentEl.ElementFromPathExist("VKr"){
agntph := AgentEl.ElementFromPath("VKr").Name
agntph := RegExReplace(agntph, "\D", "")
  }
 }
}

if agntph = ""{
try{
AgentEl := UIA.ElementFromHandle("Agent Panel - Main window")
if AgentEl.ElementFromPathExist("VKs"){
agntph := AgentEl.ElementFromPath("VKs").Name
agntph := RegExReplace(agntph, "\D", "")
  }
 }
}
; if agntph != ""
; agntph := RegExReplace(agntph, "\D", "")
; Else
; MsgBox "Could not get number from LiveVox","Error 0x1247",48+262144
} ;if livevox is alive
Else
MsgBox "Could not find LiveVox","Error 0x1250",48+262144
} ;if e3 is empty
else
agntph := noteGui["e3"].Value
agntph := RegExReplace(agntph, "\D", "")
if agntph = ""{
msgbox "Unable to retrive phone number from Livevox. Please enter a phone number and try again.","Error 0x1256",48+262144
Exit
}
if LatitudeEl.ElementFromPath("RYYIYR").Name = "Note Entry Form"
Return
else{
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
phll := LatitudeEl.ElementFromPath("RYYIYb").Children ;counts the length of the list elements (phone list length)
loopl := phll.length 	;turns the above into a variable 
; msgbox "total elements to check: " loopl
cbph := agntph 			;sets comparator to clipboard (clipboard phone (for now, and then we are going to start stealing from livevox too))
phln := "" 				;variable to assign our results to (phone list number)
nmf := 0 				;no match found conditional variable

while phln !== cbph{ 			;while the result does not equal input
yarr := A_index 	 			;counts our loops for us, which is also our row number
phln := LatitudeEl.WalkTree("16,1,1,1,1,6," A_index ",2").Value ;sets our result var to next list item
; msgbox "loop #: '" yarr "'...total list ln: '" loopl "'...phone # I'm looking for: '" cbph "'... What I am checking against: '" phln "'"
if phln = cbph{					;if we have a match
break 							;stop
} 
else if A_index >= loopl { 		;if we loop through the whole list with no match
MsgBox "Number not on list","Error",48+262144
nmf := 1 						;no match found
break
 }
}

if nmf = 0{
element1 := LatitudeEl.ElementFromPathExist("16,1,1,1,1,6,2")
if element1.name = "Vertical Scroll Bar"{
scrl1 := Abs((yarr-3)/(loopl-3))
scrl1 := Round(scrl1 * 100)
;MsgBox (yarr-3) " / " ((loopl-3)+1) " = " scrl1
if scrl1 <= 0{
scrl1 := 0
}
element1.Value := 0
element1.Value := scrl1
}
try{
LatitudeEl.ElementFromPath("16,1,1,1,1,6,1").Value := 0
sleep 100
SetControlDelay -1
LatitudeEl.ElementFromPath("16,1,1,1,1,6," yarr ",7").ControlClick()
;LatitudeEl.ElementFromPath("16,1,1,1,1,6," yarr ",7").Click()
}
catch as e
MsgBox "There was an error processing your request. If the issue persists, please restart note assist.","Error 0x" e.line,48+262144
  }
 }
}
Else
MsgBox "Latitude not found","Error0x1313",48+262144
}

/**************
*    ibob     *
***************/

; OnMessage(0x44, ibobmsg)
; global gMsgBoxTitle := "Inbound/Outbound"
; global ibobresult := MsgBox("Please select inbound or outbound:", gMsgBoxTitle, "YesNoCancel")


; ibobmsg(wParam, lParam, msg, hwnd){
    ; global gMsgBoxTitle
    ; DetectHiddenWindows true
    ; msgBoxHwnd := WinExist(gMsgBoxTitle)
    ; if (msgBoxHwnd) {
      ; ControlSetText "Inbound", "Button1", msgBoxHwnd 	;yes
      ; ControlSetText "Outbound", "Button2", msgBoxHwnd 	;no
    ; }
    ; DetectHiddenWindows false
; }

/*****************
*   SIF CALC     *
*****************/

		noteGui["Title"].GetPos(&titleX, &titleY, &titleW, &titleH)		
		noteGui.Add("Text","x" titleX+nttw+ntew+Margin " y" Margin " w" ntew+margin " h" Margin*3 " 0x200 Center Border vCalcTitle","SIF Calculator")
		noteGui["CalcTitle"].SetFont("S" Margin*1.75 " w" Margin*100)
		
		notegui["E1"].GetPos(&tempx, &tempy, &tempw, &temph)
		noteGui.Add("Text","x" titleX+nttw+ntew+Margin " y" tempy " w" nttw+50 " h" ntth " 0x200 vcalct1", "Principal Balance:")
		noteGui["CalcTitle"].GetPos(&titlecX, &titlecY, &titlecW, &titlecH)
		notegui["calct1"].GetPos(&tempx, &tempy, &tempw, &temph)
		noteGui.Add("Edit","x" tempx+nttw+50 " y" tempy-1 " w" titlecW-tempw " h" ntth " 0x200 vcalce1", "0.00").OnEvent("Change", CalcUPD)
		
		noteGui["calce1"].OnEvent("LoseFocus",adhd0)
		adhd0(Edit, Info){
		if !IsNumber(noteGui["calce1"].value){
		MsgBox "Please use numbers only.","Error",48+262144
		noteGui["calce1"].value := "0.00"
		}
		CalcUPD(1,1)
		}
		
		notegui["E2"].GetPos(&tempx, &tempy, &tempw, &temph)
		noteGui.Add("Text","x" titleX+nttw+ntew+Margin " y" tempy " w" nttw+50 " h" ntth " 0x200 vcalct2", "Discount %:")
		noteGui["CalcTitle"].GetPos(&titleX, &titleY, &titleW, &titleH)
		notegui["calct2"].GetPos(&tempx, &tempy, &tempw, &temph)
		noteGui.Add("Edit","x" tempx+nttw+50 " y" tempy-1 " w" (titleW-tempw)/2+10 " h" ntth " 0x200 vcalce2", "0").OnEvent("Change", CalcUPD)
		noteGui.Add("UpDown", "vdiscUD Range0-100","0")
		
		noteGui["calce2"].OnEvent("LoseFocus",adhd)
		adhd(Edit, Info){
		if !IsNumber(noteGui["calce2"].value){
		MsgBox "Please use numbers only.","Error",48+262144
		noteGui["calce2"].value := 0
		}
		CalcUPD(1,1)
		}
		
		notegui["cb1DOB"].GetPos(&tempx, &tempy, &tempw, &temph)
		noteGui.Add("Text","x" titleX " y" tempy-2 " w" nttw+50 " h" ntth " 0x200 vcalct3", "Discount Amount:")
		noteGui["CalcTitle"].GetPos(&titleX, &titleY, &titleW, &titleH)
		notegui["calct3"].GetPos(&tempx, &tempy, &tempw, &temph)
		noteGui.Add("Edit","x" tempx+nttw+50 " y" tempy-1 " w" (titleW-tempw) " h" ntth " 0x200 vcalce3 ReadOnly", "")
		
		If settings["esls"] = 0{ 
		notegui["e3"].GetPos(&tempx, &tempy, &tempw, &temph)
		noteGui.Add("Text","x" titleX " y" tempy " w" nttw+50 " h" ntth " 0x200 vcalct4", "Balance After Disc:")
		noteGui["CalcTitle"].GetPos(&titleX, &titleY, &titleW, &titleH)
		notegui["calct4"].GetPos(&tempx, &tempy, &tempw, &temph)
		noteGui.Add("Edit","x" tempx+nttw+50 " y" tempy-1 " w" (titleW-tempw) " h" ntth " 0x200 vcalce4 ReadOnly", "")	
		notegui["impb3"].GetPos(&tempx, &tempy, &tempw, &temph)
		noteGui.Add("Edit","x" tempx+ntew+Margin " y" tempy " w" titleW " h" ntth*3 " 0x200 vcalce5 ReadOnly", "SIF APPROVED $ " notegui["calce4"].Value " (PRIN BAL OF " notegui["calce1"].Value " @ " notegui["calce2"].Value "% DISC)")
		}
		
		If settings["esls"] = 1{ 
		notegui["cb5ls2"].GetPos(&tempx, &tempy, &tempw, &temph)
		notegui["calct3"].GetPos(&tempxt, &tempyt, &tempwt, &tempht)
		noteGui.Add("Text","x" tempxt " y" tempy " w" nttw+50 " h" ntth " 0x200 vcalct4", "Balance After Disc:")
		noteGui["CalcTitle"].GetPos(&titleX, &titleY, &titleW, &titleH)
		notegui["calct4"].GetPos(&tempx, &tempy, &tempw, &temph)
		noteGui.Add("Edit","x" tempx+nttw+50 " y" tempy-1 " w" (titleW-tempw) " h" ntth " 0x200 vcalce4 ReadOnly", "")
		notegui["e3"].GetPos(&tempx, &tempy, &tempw, &temph)
		notegui["tls3"].GetPos(&tempxt, &tempyt,,)
		noteGui.Add("Edit","x" tempx+tempw+Margin " y" tempyt " w" titleW " h" ntth*3 " 0x200 vcalce5 ReadOnly", "SIF APPROVED $ " notegui["calce4"].Value " (PRIN BAL OF " notegui["calce1"].Value " @ " notegui["calce2"].Value "% DISC)")
		}
		
		noteGui["b9"].GetPos(&rx,&ry,&rw,&b9h)
		noteGui.Add("Button","x" titleX+titleW-btnw " y" ry " w" btnw " h" btnh " vcalcb1", "Close Calc").OnEvent("Click", CalcCloseb)
		
		;notegui["calce1"].OnEvent("Change", CalcUPD)

		CalcUPD(Edit, Info){
		
		disc := 0
		pri := 0
		discbal := 0
		
		if notegui["calce1"].Value = ""
		pri := 0
		else
		if !isnumber(notegui["calce1"].Value)
		pri := 0
		else
		pri := notegui["calce1"].Value
		

		
		if notegui["calce2"].Value = ""{
		disc := 0
		}else{
		if !isnumber(notegui["calce2"].Value){
		disc := 0
		}else{
		disc := notegui["calce2"].Value
		}
		}
		
		discbal := pri*(disc/100)
		discbal := floor(discbal*100)
		discbal := discbal/100
		finalbal := pri-discbal
		
		notegui["calce3"].value := Round(discbal,2)
		notegui["calce4"].value := Round(finalbal, 2)
		notegui["calce5"].value := "SIF APPROVED $" notegui["calce4"].Value " (PRIN BAL OF $" notegui["calce1"].Value " @ " notegui["calce2"].Value "% DISC)"
		}
		

		CalcClose(){
		loop 5{
		 notegui["calce" A_index].Value 	:= ""
		}
		
		notegui["CalcTitle"].Visible 		:= 0
		loop 4{
		 notegui["calct" A_index].Visible 	:= 0
		 notegui["calce" A_index].Visible 	:= 0
		}
		notegui["calce5"].Visible 			:= 0
		notegui["discUD"].Visible 			:= 0
		notegui["calcb1"].Visible 			:= 0
		noteGui.Hide
		noteGui.Show("Autosize")
		}		
		
		CalcCloseb(Button, Info){
		Result := MsgBox("Would you like to close the SIF Calculator","Close SIF?",32+262144 " YesNo")
		if Result = "No"
		Return
		else
		
		loop 4{
		 notegui["calce" A_index].Value 	:= ""
		}
		
		notegui["CalcTitle"].Visible 		:= 0
		loop 4{
		 notegui["calct" A_index].Visible 	:= 0
		 notegui["calce" A_index].Visible 	:= 0
		}
		notegui["calcb1"].Visible 			:= 0
		noteGui.Hide
		noteGui.Show("Autosize")
		notegui["calce5"].Visible 			:= 0
		notegui["discUD"].Visible 			:= 0
		noteGui.Hide
		noteGui.Show("Autosize")
		}
		
		CalcOpen(){
		
		If noteGui["slepaTitle"].Visible = 1{
		Result := MsgBox("Would you like to close the SLEPA?","Close SLEPA?",32+262144 " YesNo")
		if Result = "No"
		Return
		Else
		CloseSLEPA()
		}		
		
		If noteGui["EPATitle"].Visible = 1{
		Result := MsgBox("Would you like to close the EPA?","Close?",32+262144 " YesNo")
		if Result = "No"
			Return
		else
		CloseEPA()
		}

		If noteGui["FCTitle"].Visible = 1{
		Result := MsgBox("Would you like to close the F/C?","Close?",32+262144 " YesNo")
		if Result = "No"
			Return
		else		
		CloseFC()
		}		

		If noteGui["atytitle"].Visible = 1{
		Result := MsgBox("Would you like to close the Attorney note","Close aty?",32+262144 " YesNo")
		if Result = "No"
		Return
		else
		CloseATY()
		}
		
		For bn, label in noteBSCMap{
		noteGui["sb" bn].Visible := 0
		}
		noteGui["gb1"].Visible := 0
		noteGui["gb2"].Visible := 0
		
		noteGui["CalcTitle"].Visible 		:= 1
		loop 4{
		 notegui["calct" A_index].Visible 	:= 1
		 notegui["calce" A_index].Visible 	:= 1
		}
		notegui["calce5"].Visible 			:= 1
		notegui["discUD"].Visible 			:= 1
		notegui["calcb1"].Visible 			:= 1
		
		CalcUPD(1,1)
		noteGui.Hide
		noteGui.Show("Autosize")
		}
		
/*************************
*	 Last Row Checks	 *
*************************/
lastrowcheck(){
if notegui["cbdisp"].value = "x"{
if LatitudeEl.ElementFromPath("RYYIYR2/").ToggleState = "1"{
notegui["cbdisp"].value := notegui["cbdisp"].value
}
else{
LatitudeEl.ElementFromPath("RYYIYR2/").Toggle()
 }
}

if notegui["cbcease"].value = "x"{
if LatitudeEl.ElementFromPath("RYYIYR2r").ToggleState = "1"{
notegui["cbcease"].value := notegui["cbcease"].value
}
else{
LatitudeEl.ElementFromPath("RYYIYR2r").Toggle()
 }
}

if notegui["cbbad"].value = "x"{
if LatitudeEl.ElementFromPath("RYYIYR2q").ToggleState = "1"{
notegui["cbbad"].value := notegui["cbbad"].value
}
else{
LatitudeEl.ElementFromPath("RYYIYR2q").Toggle()
 }
}

if notegui["cbconsent"].value = "x"{
if LatitudeEl.ElementFromPath("RYYIYR2").ToggleState = "1"{
notegui["cbconsent"].value := notegui["cbconsent"].value
}
else{
LatitudeEl.ElementFromPath("RYYIYR2").Toggle()
  }
 }
}

/*****************************
*	 PIFLR for all accounts	 *
*****************************/
piflrall(){
if WinExist("Latitude - Account Work Form"){
If noteGui["e4"].Text = ""
MsgBox "Fill in email and try again"
else{
Result := MsgBox("This action will post a PIFLR note to each file. Once started, this process can not be interuptted. Are you sure you would like to continue?","Auto PIFLR",48+262144 " YesNo")
if Result = "No"
Return
else
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
WinGetPos &latx, &laty,,, "Latitude - Account Work Form"
LatitudeEl.ElementFromPath("Yw0").Invoke()
sleep 250
SetTitleMatchMode "RegEx"
WinActivate "Link \d{4,8}"
WinMove latx, laty,,, "Link \d{4,8}"
LinkEl := UIA.ElementFromHandle("Link \d{4,8}")
numfiles := LinkEl.ElementFromPath("YYqNO").Children

loop numfiles.length{
LinkEl.WalkTree("1,2,1,1," A_index).Select()
LinkEl.ElementFromPath("1,1,1,1").Select()
;pause
datev := UIA.ElementFromHandle("Link \d{4,8}")
if datev.ElementFromPath("1,6,1,1,6").Value = ""{
LinkEl.ElementFromPath("1,1,1,3").Select()
		SendInput "!n"
		WinWait "Latitude - New Note"
		If WinExist("Latitude - New Note"){
		A_clipboard := noteGui["e4"].Text
		WinActivate "Latitude - New Note"
		SendInput "CO{Down}{Tab}PIFLR{Down}{Tab}"
		SendInput "^v"
		SendInput "!a"
		sleep 100
	}
	LinkEl.ElementFromPath("1,1,1,1").Select()
   }
  If A_index = numfiles.length
  MsgBox "Done","Done",64+262144
  }
 }
}
Else
MsgBox "Latitude not found","Error",48+262144
}

/*************************************
*	 Remove Cease for all accounts	 *
*************************************/
ceaseremall(){ ;cease, remove all
if WinExist("Latitude - Account Work Form"){
Result := MsgBox("This action will remove the Cease Comm on each file. Once started, this process can not be interuptted. Are you sure you would like to continue?","Remove Cease",48+262144 " YesNo")
if Result = "No"
Return
else
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
WinGetPos &latx, &laty,,, "Latitude - Account Work Form"
LatitudeEl.ElementFromPath("Yw0").Invoke()
sleep 250
SetTitleMatchMode "RegEx"
WinActivate "Link \d{4,8}"
WinMove latx, laty,,, "Link \d{4,8}"
LinkEl := UIA.ElementFromHandle("Link \d{4,8}")
numfiles := LinkEl.ElementFromPath("YYqNO").Children

loop numfiles.length{
LinkEl.ElementFromPath("1,1,1,1").Select()
; MsgBox A_index
LinkEl.WalkTree("1,2,1,1," A_index).Select()
LinkEl.ElementFromPath("1,1,1,1").Select()
;pause
datev := UIA.ElementFromHandle("Link \d{4,8}")
if datev.ElementFromPath("1,6,1,1,6").Value = ""{
LinkEl.ElementFromPath("1,1,1,2").Select()
hotnote := LinkEl.ElementFromPath("1,6,1,2").Value
hotnote := RegExReplace(hotnote,"(?i)\s*cease\s+comm(UNICATION)?\s*", " ")
; hotnote := RegExReplace(hotnote,"(?i)\s*cease\s+", " ")
; hotnote := RegExReplace(hotnote,"(?i)\s*cease?\s*", " ")
LinkEl.ElementFromPath("1,6,1,2").Value := hotnote
LinkEl.FindElement({Type: "Button", Name: "More Info"}).Invoke()
WinWait "Debtor Details"
WinActivate "Debtor Details"
SendInput "!r"
sleep 250
LatitudeEl := UIA.ElementFromHandle("Debtor Details")
if LatitudeEl.ElementFromPath("1,2,1,4").ToggleState = 1
LatitudeEl.ElementFromPath("1,2,1,4").ToggleState   := 0
if LatitudeEl.ElementFromPath("1,2,1,5").ToggleState = 1
LatitudeEl.ElementFromPath("1,2,1,5").ToggleState   := 0
if LatitudeEl.ElementFromPath("1,2,1,8").ToggleState = 1
LatitudeEl.ElementFromPath("1,2,1,8").ToggleState   := 0
; LatitudeEl.ElementFromPath(["1,2,1,4", "1,2,1,5", "1,2,1,8"])

; if LatitudeEl.ElementFromPath("YY0").IsEnabled = 1{
; LatitudeEl.ElementFromPath("YY0").Invoke()
; while LatitudeEl.ElementFromPath("YY0").IsEnabled = 1{
; sleep 10
; if LatitudeEl.ElementFromPath("YY0").IsEnabled = 0{
; LatitudeEl.ElementFromPath("1,1,2").Invoke()
; ; if WinExist "Debtor Details"
; ; WinClose "Debtor Details"
; sleep 250
; break
   ; }
  ; }
 ; }
; Else{
LatitudeEl.ElementFromPath("1,1,2").Invoke()
; MsgBox "going back"
 ; }
}
  If A_index = numfiles.length{
  LinkEl.ElementFromPath("1,1,1,1").Select()
  WinClose "Link \d{4,8}"
  SetTitleMatchMode 2
  MsgBox "Done","Done",64+262144
   }
  }
 }
Else
MsgBox "Latitude not found","Error",48+262144
}

/*************************
*	Cease all accounts	 *
*************************/
ceaseall(){ ;cease, all files
if WinExist("Latitude - Account Work Form"){
Result := MsgBox("This action will add a Cease Comm on each file. Once started, this process can not be interuptted. Are you sure you would like to continue?","Cease Comm Remove",48+262144 " YesNo")
if Result = "No"
Return
else
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
WinGetPos &latx, &laty,,, "Latitude - Account Work Form"
LatitudeEl.ElementFromPath("Yw0").Invoke()
sleep 250
SetTitleMatchMode "RegEx"
WinActivate "Link \d{4,8}"
WinMove latx, laty,,, "Link \d{4,8}"
LinkEl := UIA.ElementFromHandle("Link \d{4,8}")
numfiles := LinkEl.ElementFromPath("YYqNO").Children

loop numfiles.length{
LinkEl.WalkTree("1,2,1,1," A_index).Select()
LinkEl.ElementFromPath("1,1,1,1").Select()
;pause
datev := UIA.ElementFromHandle("Link \d{4,8}")
if datev.ElementFromPath("1,6,1,1,6").Value = ""{
LinkEl.ElementFromPath("1,1,1,2").Select()
hotnote := LinkEl.ElementFromPath("1,6,1,2").Value
hotnotev := hotnote
hotnotev := RegExMatch(hotnotev,"(?i)\s*cease\s+comm(UNICATION)?\s*")
if hotnotev != 0
hotnotev := hotnotev
else{
if hotnote = ""
hotnote .= "CEASE COMM"
else
hotnote .= "`r`nCEASE COMM"
}
sleep 100
LinkEl.ElementFromPath("1,6,1,2").Value := hotnote
LinkEl.FindElement({Type: "Button", Name: "More Info"}).Invoke()
WinWait "Debtor Details"
WinActivate "Debtor Details"
SendInput "!r"
sleep 250
LatitudeEl := UIA.ElementFromHandle("Debtor Details")
; if LatitudeEl.ElementFromPath("1,2,1,4").ToggleState = 1
; LatitudeEl.ElementFromPath("1,2,1,4").ToggleState   := 0
; if LatitudeEl.ElementFromPath("1,2,1,5").ToggleState = 1
; LatitudeEl.ElementFromPath("1,2,1,5").ToggleState   := 0
if LatitudeEl.ElementFromPath("1,2,1,8").ToggleState = 0
LatitudeEl.ElementFromPath("1,2,1,8").ToggleState   := 1
; LatitudeEl.ElementFromPath(["1,2,1,4", "1,2,1,5", "1,2,1,8"])

; if LatitudeEl.ElementFromPath("YY0").IsEnabled = 1{
; LatitudeEl.ElementFromPath("YY0").Invoke()
; while LatitudeEl.ElementFromPath("YY0").IsEnabled = 1{
; if LatitudeEl.ElementFromPath("YY0").IsEnabled = 0{
; LatitudeEl.ElementFromPath("1,1,2").Invoke()
; WinClose "Debtor Details"
; sleep 250
   ; }
  ; }
 ; }
; Else
LatitudeEl.ElementFromPath("1,1,2").Invoke()
; WinClose "Debtor Details"
; sleep 250
}
  ; MsgBox "A_index = " A_index " and numfiles = " numfiles.length
  If A_index = numfiles.length{
  LinkEl.ElementFromPath("1,1,1,1").Select()
  WinClose "Link \d{4,8}"
  SetTitleMatchMode 2
  MsgBox "Done","Done",64+262144
   }
  }
 }
Else
MsgBox "Latitude not found","Error",48+262144
}

callerror(){
Result1 := "Unable to process request, please try again. If the error persists, please restart NoteAssist.","Error",5+16+242144
if Result1 = "Cancel"
return
else
bclick3(1,1)
}

/************
* CANTC ALL *
************/

cantcall(){
cantcnote := "SEND CONSUMER THE FOLLOWING: ITEMIZED STATEMENT"
if WinExist("Latitude - Account Work Form"){
Result := MsgBox("This action will post a CO/CANTC note to each file. Once started, this process can not be interuptted. Are you sure you would like to continue?","Auto CANTC",48+262144 " YesNo")
if Result = "No"
Return
else
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
WinGetPos &latx, &laty,,, "Latitude - Account Work Form"
LatitudeEl.ElementFromPath("Yw0").Invoke()
sleep 250
SetTitleMatchMode "RegEx"
WinActivate "Link \d{4,8}"
WinMove latx, laty,,, "Link \d{4,8}"
LinkEl := UIA.ElementFromHandle("Link \d{4,8}")
numfiles := LinkEl.ElementFromPath("YYqNO").Children

loop numfiles.length{
LinkEl.WalkTree("1,2,1,1," A_index).Select()
LinkEl.ElementFromPath("1,1,1,1").Select()
;pause
datev := UIA.ElementFromHandle("Link \d{4,8}")
if datev.ElementFromPath("1,6,1,1,6").Value = ""{
LinkEl.ElementFromPath("1,1,1,3").Select()
		SendInput "!n"
		WinWait "Latitude - New Note"
		If WinExist("Latitude - New Note"){
		A_clipboard := "SEND CONSUMER THE FOLLOWING: ITEMIZED STATEMENT"
		WinActivate "Latitude - New Note"
		SendInput "CO{Down}{Tab}CANTC{Down}{Tab}"
		SendInput "+{END}" "^v"
		SendInput "!a"
		sleep 100
	}
	LinkEl.ElementFromPath("1,1,1,1").Select()
   }
  If A_index = numfiles.length
  MsgBox "Done","Done",64+262144
  }
 }
Else
MsgBox "Latitude not found","Error",48+262144
}

/*************
*    PTPIF   *
*************/

ptpif(){
if WinExist("PTPDays")
winclose("PTPDays")
ptpifgui := gui("AlwaysOnTop ToolWindow","PTPDays")
ptpifgui.Add("Text","vptpt","Payment in [#] of days:")
ptpifgui.Add("Edit","w50 vptpife center 0x200","").OnEvent("Change", DateChange)
ptpifgui.Add("UpDown","vptpud Range0-30","")
ptpifgui.Add("Button","vok","Ok").OnEvent("Click", ptpifok)
daysv := ptpifgui["ptpife"].Value
ptpday := DateAdd(A_now, daysv, "days")
ptpday := FormatTime(ptpday,"M/d/yyyy")

ptpifok(Button, Info){
noteGui.opt("-AlwaysOnTop")
if ptpifgui["ptpife"].Value > 30
ptpifgui["ptpife"].Value := 30
if ptpifgui["ptpife"].Value < 0
ptpifgui["ptpife"].Value := 0
daysv := ptpifgui["ptpife"].Value
ptpday := DateAdd(A_now, daysv, "days")
ptpday := FormatTime(ptpday,"M/d/yyyy")
Result := MsgBox("This will put a PIF promise on account. Are you sure you would like to continue?","PTPIF","YesNo " 32+262144)
if Result = "No"
Return
Else
WinClose "Days ahk_exe AutoHotkey64.exe"
if WinExist("Latitude - Account Work Form"){
try{
goodbyepops()
if WinExist("Enter New Arrangements"){
WinClose("Enter New Arrangements")
}
WinActivate("Latitude - Account Work Form")
SendInput "!p"
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
if LatitudeEl.ElementExist({Type: "50000", Name: "Delete Arrangement", LocalizedType: "button"}){
Result := MsgBox("There is already an arrangement on the account. Would you like to delete the arrangement?","PTPIF","YesNo " 32+262144)
if Result = "Yes"{
try{
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
LatitudeEl.FindElement({Type: "50000", Name: "Delete Arrangement", LocalizedType: "button"}).ControlClick()
WinWait("Modify Payments ahk_exe Latitude.exe")
ControlClick("OK", "Modify Payments ahk_exe Latitude.exe")
; LatitudeEl := UIA.ElementFromHandle("Modify Payments ahk_exe Latitude.exe")
; LatitudeEl.ElementFromPath("Y0/").ControlClick()
WinWait("Arrangements Panel ahk_exe Latitude.exe")
LatitudeEl := UIA.ElementFromHandle("Arrangements Panel ahk_exe Latitude.exe")
LatitudeEl.ElementFromPath("0").ControlClick()
}
catch as e{
msgbox e.message " and " e.line
}
}
Else
Return
}
; ControlClick [Control-or-Pos, WinTitle, WinText, WhichButton, ClickCount, Options, ExcludeTitle, ExcludeText]
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form ahk_exe Latitude.exe")
LatitudeEl.FindElement({Type: "50000", Name: "Negotiate Arrangement", LocalizedType: "button"}).ControlClick()
WinWait("Enter New Arrangements")
WinActivate("Enter New Arrangements")
WinWaitActive("Enter New Arrangements")
LatitudeEn := UIA.ElementFromHandle("Enter New Arrangements")
IF LatitudeEn.ElementExist({Type: "50005", Name: "Select Accounts", Value: "Select Accounts", LocalizedType: "link"}){
LatitudeEn.FindElement({Type: "50005", Name: "Select Accounts", Value: "Select Accounts", LocalizedType: "link"}).ControlClick()
LatitudeEn.WaitElement({Type: "50000", Name: "Select All", LocalizedType: "button", AutomationId: "selectAllButton"}).ControlClick()
}
ELSE
sleep 100
LatitudeEn := UIA.ElementFromHandle("Enter New Arrangements")
;MsgBox "now"
tiap := LatitudeEn.FindElement({Type: "50005", Name: "Calculate Payments", Value: "Calculate Payments", LocalizedType: "link"})
tiap.Invoke()
LatitudeEn := UIA.ElementFromHandle("Enter New Arrangements")
sleep 500
ptparr := StrSplit(ptpday,"/")
LatitudeEn := UIA.ElementFromHandle("Enter New Arrangements")
LatitudeEn.ElementFromPathExist("YYYYYY").SetFocus()
SendInput ptparr[1] "{Right}" ptparr[2] "{Right}" ptparr[3]
LatitudeEn.FindElement({Type: "50005", Name: "Review Arrangement", Value: "Review Arrangement", LocalizedType: "link"}).Invoke()
LatitudeEn := UIA.ElementFromHandle("Enter New Arrangements")
LatitudeEn.ElementFromPath("Y0q").Invoke()
WinWait("Writing Payments")
WinWaitClose("Writing Payments")
MsgBox "Pay In Full Promise set for " ptpday ".","Success",64+262144
} ;try
catch as e{
MsgBox "Unable to process request.","Error 0x" e.line,16+262144
} ;catch
} ;if WinExist
Else
MsgBox "Latitude not found","Error",48+262144
noteGui.opt("AlwaysOnTop")
} ;ptpifok

ptpifgui.add("Button","yp vclose","Close").OnEvent("Click", ptpifclose)
ptpifclose(Button, Info){
WinClose "Days ahk_exe AutoHotkey64.exe"
}

ptpifgui["ptpt"].GetPos(&tx,&ty,&tw,&th)
ptpifgui["ptpife"].move(,,tw)
ptpifgui["ptpife"].GetPos(&ex,&ey,&ew,&eh)
ptpifgui["ptpud"].GetPos(&ux,&uy,&uw,&uh)
ptpifgui["ptpife"].move(,,ew-uw,eh-5)
ptpifgui["ptpud"].move(ex+ew-uw,,,uh-5)
ptpifgui["close"].GetPos(&bx,&by,&bw,&bh)
ptpifgui["close"].move(ex+ew-bw,by-margin)
ptpifgui["ok"].move(,by-margin)
ptpifgui["ok"].GetPos(&bx,&by,&bw,&bh)
ptpifgui.Add("Text","x" bx " y" by+bh+margin/2.5 " h" nteh " vdatedt 0x200","Date: ")
ptpifgui.Add("text","yp w" ew-margin*3.8 " h" nteh " vdatede 0x200 right",ptpday)

DateChange(Edit, Info){
daysv := ptpifgui["ptpife"].Value
ptpday := DateAdd(A_now, daysv, "days")
ptpday := FormatTime(ptpday,"M/d/yyyy")
ptpifgui["datede"].value := ptpday
ptpifgui["datede"].redraw
}

ptpifgui.show("Autosize")

}

/*************************
* Phone Look Up Function *
*************************/

phonelookupfun(){
if noteGui["e3"].Value = ""{
if WinExist("ahk_exe Agent Desktop Native.exe"){
try{
;MsgBox "try1 phone"
AgentEl := UIA.ElementFromHandle("ahk_exe Agent Desktop Native.exe")
if AgentEl.ElementFromPathExist("VKr"){
agntph := AgentEl.ElementFromPath("VKr").Name
agntph := RegExReplace(agntph, "\D", "")
 }
}

if agntph = ""{
try{
AgentEl := UIA.ElementFromHandle("ahk_exe Agent Desktop Native.exe")
if AgentEl.ElementFromPathExist("VKs"){
agntph := AgentEl.ElementFromPath("VKs").Name
agntph := RegExReplace(agntph, "\D", "")
  }
 }
}

if agntph = ""{		
try{
;MsgBox "try1 phone"
AgentEl := UIA.ElementFromHandle("Agent Panel - Main window")
if AgentEl.ElementFromPathExist("VKr"){
agntph := AgentEl.ElementFromPath("VKr").Name
agntph := RegExReplace(agntph, "\D", "")
  }
 }
}

if agntph = ""{
try{
AgentEl := UIA.ElementFromHandle("Agent Panel - Main window")
if AgentEl.ElementFromPathExist("VKs"){
agntph := AgentEl.ElementFromPath("VKs").Name
agntph := RegExReplace(agntph, "\D", "")
  }
 }
}
; if agntph != ""
; agntph := RegExReplace(agntph, "\D", "")
; Else
; MsgBox "Could not get number from LiveVox","Error 0x1247",48+262144
} ;if livevox is alive
Else
MsgBox "Could not find LiveVox","Error 0x1250",48+262144
} ;if e3 is empty
else
agntph := noteGui["e3"].Value
agntph := RegExReplace(agntph, "\D", "")
if agntph = ""{
msgbox "Unable to retrive phone number from Livevox. Please enter a phone number and try again.","Error 0x1256",48+262144
Exit
}
}

/*****************************
*	    Button Functions	 *
*****************************/

/*************
*  BUTTON 1  *
*************/

bclick1(Button,Info){ ;Clears the controls
	b1str := ""
	A_clipboard := ""
	Result := MsgBox("Would you like to clear the form?","Clear?",32+262144 " YesNo")
		if Result = "No"
			Return
		else
	For val, etext in notectrl{
		if noteGui["e" val].Text = ""
			A_clipboard := A_clipboard
		else
			A_clipboard := A_clipboard etext " " noteGui["e" val].Text delim
			,noteGui["e" val].Text := ""
		
		if A_Index = 2{
			For val, cbtext in noteCBm{ ;leaving it this way because I have gaslit myself into thinking that it might have some use
				if settings["esls"] = 0{
					if noteGui[val].Value = ""
						A_clipboard := A_clipboard 
					else
						A_clipboard := A_clipboard cbtext delim noteGui[val].Text := ""
					if A_index = 3
						Break
				}
				
				Else if settings["esls"] = 1{
				if noteGui[val].Value = ""
					A_clipboard := A_clipboard 
				else
					A_clipboard := A_clipboard cbtext delim noteGui[val].Text := ""
				if A_index = 3
					Break
				}
			} ;End CB For Loop					
		} ;End IF INDEX 2
		
		if A_index = 4{
		
			if settings["esls"] = 1{
				if noteGui["cb4ls1"].Text = "" ;int to req jmt
				A_clipboard := A_Clipboard
				else
				A_clipboard := A_Clipboard noteCBm["cb4ls1"] delim noteGui["cb4ls1"].Text := ""	
			}
		}
		
		if A_index = 6{
		
			if settings["esls"] = 1{	
			
				if noteGui["els3"].Value = "" ;adv balance w/ fees
				A_clipboard := A_Clipboard
				else
				A_clipboard := A_clipboard "ADV BAL W/ FEES: " noteGui["els3"].Value "" delim
				noteGui["els3"].Value := ""
			}
		}
		
		if A_index = 7{
		
			if settings["esls"] = 1{

			; if noteGui["cb4ls1"].Text = "" ;int to req jmt
			; A_clipboard := A_Clipboard
			; else
			; A_clipboard := A_Clipboard noteCBm["cb4ls1"] delim noteGui["cb4ls1"].Text := ""
		
			if noteGui["els3"].Value = "" ;adv balance w/ fees
			A_clipboard := A_Clipboard
			else
			
			if noteGui["cb5ls2"].Text = "" ;will not act
			A_clipboard := A_Clipboard
			else
			A_clipboard := A_Clipboard noteCBm["cb5ls2"] delim noteGui["cb5ls2"].Text := ""
			}
		}		
	} ;end for loop

	noteGui["cbdisp"].Text 		:= ""
	noteGui["cbcease"].Text 	:= ""
	noteGui["cbbad"].Text 		:= ""
	noteGui["cbconsent"].Text  	:= ""
	
	A_Clipboard := RegExReplace(A_Clipboard, "`r`n", "") 	; Remove line breaks
	A_Clipboard := StrReplace(A_Clipboard, "`r", "") 		; Remove carriage returns
	A_Clipboard := StrReplace(A_Clipboard, "`t", "") 		; Remove tabs
	A_Clipboard := StrUpper(A_Clipboard) 					; Convert to uppercase				
	A_Clipboard := RTrim(A_Clipboard ," - ")				; Trims the last delim

} ;End bclick1

/*************
*  BUTTON 2  *
*************/
scbval := 0

bclick2(Button,Info){ 

If scbval = 0
scbval := 1
Else
scbval := 0

epahas := 0

If noteGui["FCTitle"].Visible = 1{
Result := MsgBox("Would you like to close the F/C?","Close F/C?",32+262144 " YesNo")
if Result = "No"
	Return
else
CloseFC()
}

If noteGui["slepaTitle"].Visible = 1{
Result := MsgBox("Would you like to close the SLEPA?","Close SLEPA?",32+262144 " YesNo")
if Result = "No"
Return
Else
CloseSLEPA()
}

If noteGui["EPATitle"].Visible = 1{
	Result := MsgBox("Would you like to close the EPA?","Close EPA?",32+262144 " YesNo")
	if Result = "No"
		Return
	else
	CloseEPA()
}

If noteGui["CalcTitle"].Visible = 1{
Result := MsgBox("Would you like to close the SIF Calculator","Close SIF?",32+262144 " YesNo")
if Result = "No"
Return
else
CalcClose()
}

If noteGui["atytitle"].Visible = 1{
Result := MsgBox("Would you like to close the Attorney note","Close aty?",32+262144 " YesNo")
if Result = "No"
Return
else
CloseATY()
}

For bn, label in noteBSCMap{
noteGui["sb" bn].Visible := scbval
}

noteGui["gb1"].Visible := scbval
noteGui["gb2"].Visible := scbval

noteGui.Hide
noteGui.Show("Autosize")
}

/*************
*  BUTTON 3  *
*************/

bclick3(Button,Info){ ;Copys the values from the controls
b3str := ""
A_clipboard := ""

	For val, etext in notectrl{
		if noteGui["e" val].Text = ""
			A_clipboard := A_clipboard
		else
			A_clipboard := A_clipboard etext " " noteGui["e" val].Text delim
		
		if A_Index = 2{
			For val, cbtext in noteCBm{ ;leaving it this way because I have gaslit myself into thinking that it might have some use
				if settings["esls"] = 0{
					if noteGui[val].Value = ""
						A_clipboard := A_clipboard 
					else
						A_clipboard := A_clipboard cbtext delim
					if A_index = 3
						Break
				}
				
				Else if settings["esls"] = 1{
				if noteGui[val].Value = ""
					A_clipboard := A_clipboard 
				else
					A_clipboard := A_clipboard cbtext delim
				if A_index = 3
					Break
				}
			} ;End CB For Loop					
		} ;End IF INDEX 2

		if A_index = 4{
		
			if settings["esls"] = 1{
				if noteGui["cb4ls1"].Text = "" ;int to req jmt
				A_clipboard := A_Clipboard
				else
				A_clipboard := A_Clipboard noteCBm["cb4ls1"] delim
			}
		}


		if A_index = 6{
		
			if settings["esls"] = 1{
			
				if noteGui["els3"].Value = "" ;adv balance w/ fees
				A_clipboard := A_Clipboard
				else
				A_clipboard := A_clipboard "ADV BAL W/ FEES: " noteGui["els3"].Value "" delim
			}
		}
		
		if A_index = 7{
		
			if settings["esls"] = 1{
			
			if noteGui["cb5ls2"].Text = "" ;will not act
			A_clipboard := A_Clipboard
			else
			A_clipboard := A_Clipboard noteCBm["cb5ls2"] delim 
			}
		}		
	} ;end for loop

A_Clipboard := RegExReplace(A_Clipboard, "`r`n", "") 	; Remove line breaks
A_Clipboard := StrReplace(A_Clipboard, "`r", "") 		; Remove carriage returns
A_Clipboard := StrReplace(A_Clipboard, "`t", "") 		; Remove tabs
A_Clipboard := StrUpper(A_Clipboard) 					; Convert to uppercase					
A_Clipboard := RTrim(A_Clipboard ," - ")

} ;End bclick3

/*************
*  BUTTON 4  *
*************/

bclick4(Button,Info){ ;Inbound

bclick3(1,1)
goodbyepops()
if WinExist("Latitude - Account Work Form"){
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
		if settings["esls"] = 0{
			If WinExist("Latitude - Account Work Form"){
				WinActivate "Latitude - Account Work Form"
				SendInput "!iPPP{ENTER}"
				sleep 200
				
				try{
				LatitudeEl.ElementFromPathExist("RYYIYR/D/").ControlClick()
				}
				catch{
				Result := MsgBox("Unable to process request, please try again. If the error persists, please restart NoteAssist.","Error",5+16+8192)
				if Result = "Retry"
				bclick4(1,1)
				Exit
				}
				
				PhnSrch()
				try{
				if LatitudeEl.ElementFromPathExist("RYYIYR3/K"){
				if LatitudeEl.ElementFromPath("RYYIYR3/K").Type = 50020{
				if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "TT - TALK TO"
				LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "TT - TALK TO"
				   }
				  }
				 }
				lastrowcheck()
				}
				else
				; If not WinExist("Latitude - Account Work Form")	
				MsgBox "Latitude not found","Error",48+262144
		}
		
		else if settings["esls"] = 1{
			If WinExist("Latitude - Account Work Form"){
				WinActivate "Latitude - Account Work Form"
				SendInput "!iPPPP{ENTER}"
				sleep 200
				try{
				LatitudeEl.ElementFromPathExist("RYYIYR/D/").ControlClick()
				}
				catch{
				Result := MsgBox("Unable to process request, please try again. If the error persists, please restart NoteAssist.","Error",5+16+8192)
				if Result = "Retry"
				bclick4(1,1)
				Exit
				}
				PhnSrch()
				try{
				if LatitudeEl.ElementFromPathExist("RYYIYR3/K"){
				if LatitudeEl.ElementFromPath("RYYIYR3/K").Type = 50020{
				if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "TT - TALK TO"
				LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "TT - TALK TO"
				   }
				  }
				 }
				lastrowcheck()
				}
				else
				MsgBox "Latitude not found","Error",48+262144
		}
		
 }
Else
MsgBox "Latitude not found","Error",48+262144
}

inbcall(){
goodbyepops()
if WinExist("Latitude - Account Work Form"){
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
		if settings["esls"] = 0{
			If WinExist("Latitude - Account Work Form"){
				WinActivate "Latitude - Account Work Form"
				SendInput "!iPPP{ENTER}"
				sleep 200
				try{
				LatitudeEl.ElementFromPathExist("RYYIYR/D/").ControlClick()
				}
				catch{
				Result := MsgBox("Unable to process request, please try again. If the error persists, please restart NoteAssist.","Error",5+16+8192)
				if Result = "Retry"
				inbcall()
				Exit
				}
				PhnSrch()
				}
				else
				; If not WinExist("Latitude - Account Work Form")	
				MsgBox "Latitude not found","Error",48+262144
		}
		
		else if settings["esls"] = 1{
			If WinExist("Latitude - Account Work Form"){
				WinActivate "Latitude - Account Work Form"
				SendInput "!iPPPP{ENTER}"
				sleep 200
				try{
				LatitudeEl.ElementFromPathExist("RYYIYR/D/").ControlClick()
				}
				catch{
				Result := MsgBox("Unable to process request, please try again. If the error persists, please restart NoteAssist.","Error",5+16+8192)
				if Result = "Retry"
				inbcall()
				Exit
				}
				PhnSrch()
				}
				else
				; If not WinExist("Latitude - Account Work Form")	
				MsgBox "Latitude not found","Error",48+262144
		}

 }
Else
MsgBox "Latitude not found","Error",48+262144
}


/*************
*  BUTTON 5  *
*************/

bclick5(Button,Info){ ;Outbound

bclick3(1,1)
goodbyepops()
		if settings["esls"] = 0{
			If WinExist("Latitude - Account Work Form"){
				WinActivate "Latitude - Account Work Form"
				SendInput "!iPPP{ENTER}"
				sleep 200
				try{
				LatitudeEl.ElementFromPathExist("RYYIYR/Dq").ControlClick()
				}
				catch{
				Result := MsgBox("Unable to process request, please try again. If the error persists, please restart NoteAssist.","Error",5+16+8192)
				if Result = "Retry"
				bclick5(1,1)
				Exit
				}
				PhnSrch()
				try{
				if LatitudeEl.ElementFromPathExist("RYYIYR3/K"){
				if LatitudeEl.ElementFromPath("RYYIYR3/K").Type = 50020{
				if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "TT - TALK TO"
				LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "TT - TALK TO"
				   }
				  }
				 }
				lastrowcheck()
				}
				else
				; If not WinExist("Latitude - Account Work Form")	
				MsgBox "Latitude not found","Error",48+262144
		}
		
		else if settings["esls"] = 1{
			If WinExist("Latitude - Account Work Form"){
			WinActivate "Latitude - Account Work Form"
			SendInput "!iPPPP{ENTER}"
			sleep 200
			try{
			LatitudeEl.ElementFromPathExist("RYYIYR/Dq").ControlClick()
			}
			catch{
			Result := MsgBox("Unable to process request, please try again. If the error persists, please restart NoteAssist.","Error",5+16+8192)
			if Result = "Retry"
			bclick5(1,1)
			Exit
			}
			PhnSrch()
				try{
				if LatitudeEl.ElementFromPathExist("RYYIYR3/K"){
				if LatitudeEl.ElementFromPath("RYYIYR3/K").Type = 50020{
				if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "TT - TALK TO"
				LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "TT - TALK TO"
				   }
				  }
				 }
				lastrowcheck()
				}
			; If not WinExist("Latitude - Account Work Form")
			else
			MsgBox "Latitude not found","Error",48+262144
		}
		
}

obcall(){
goodbyepops()
		if settings["esls"] = 0{
			If WinExist("Latitude - Account Work Form"){
				WinActivate "Latitude - Account Work Form"
				SendInput "!iPPP{ENTER}"
				sleep 200
				try{
				LatitudeEl.ElementFromPathExist("RYYIYR/Dq").ControlClick()
				}
				catch{
				Result := MsgBox("Unable to process request, please try again. If the error persists, please restart NoteAssist.","Error",5+16+8192)
				if Result = "Retry"
				obcall()
				Exit
				}
				PhnSrch()
				}
				else
				; If not WinExist("Latitude - Account Work Form")	
				MsgBox "Latitude not found","Error",48+262144
		}
		
		else if settings["esls"] = 1{
			If WinExist("Latitude - Account Work Form"){
			WinActivate "Latitude - Account Work Form"
			SendInput "!iPPPP{ENTER}"
			sleep 200
			try{
			LatitudeEl.ElementFromPathExist("RYYIYR/Dq").ControlClick()
			}
			catch{
			Result := MsgBox("Unable to process request, please try again. If the error persists, please restart NoteAssist.","Error",5+16+8192)
			if Result = "Retry"
			obcall()
			Exit
			}
			PhnSrch()
			}
			; If not WinExist("Latitude - Account Work Form")
			else
			MsgBox "Latitude not found","Error",48+262144
		}
}


/*************
*  BUTTON 6  *
*************/

bclick6(Button,Info){ ;Manual

bclick3(1,1)
goodbyepops()
		if settings["esls"] = 0{
			If WinExist("Latitude - Account Work Form"){
				WinActivate "Latitude - Account Work Form"
				SendInput "!iPPP{ENTER}"
				sleep 200
				try{
				LatitudeEl.ElementFromPathExist("RYYIYR/D").ControlClick()
				}
				catch{
				Result := MsgBox("Unable to process request, please try again. If the error persists, please restart NoteAssist.","Error",5+16+8192)
				if Result = "Retry"
				bclick6(1,1)
				Exit
				}
				PhnSrch()
				try{
				if LatitudeEl.ElementFromPathExist("RYYIYR3/K"){
				if LatitudeEl.ElementFromPath("RYYIYR3/K").Type = 50020{
				if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "TT - TALK TO"
				LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "TT - TALK TO"
				   }
				  }
				 }
				lastrowcheck()
				}
			else	
				MsgBox "Latitude not found","Error",48+262144
		}
		
		else if settings["esls"] = 1{
			If WinExist("Latitude - Account Work Form"){
				WinActivate "Latitude - Account Work Form"
				SendInput "!iPPPP{ENTER}"
				sleep 200
				try{
				LatitudeEl.ElementFromPathExist("RYYIYR/D").ControlClick()
				}
				catch{
				Result := MsgBox("Unable to process request, please try again. If the error persists, please restart NoteAssist.","Error",5+16+8192)
				if Result = "Retry"
				bclick6(1,1)
				Exit
				}
				PhnSrch()
				lastrowcheck()
				try{
				if LatitudeEl.ElementFromPathExist("RYYIYR3/K"){
				if LatitudeEl.ElementFromPath("RYYIYR3/K").Type = 50020{
				if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "TT - TALK TO"
				LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "TT - TALK TO"
				   }
				  }
				 }
				lastrowcheck()
				}
			else	
				MsgBox "Latitude not found","Error",48+262144
		}
}			

mancall(){
goodbyepops()
		if settings["esls"] = 0{
			If WinExist("Latitude - Account Work Form"){
				WinActivate "Latitude - Account Work Form"
				SendInput "!iPPP{ENTER}"
				sleep 200
				try{
				LatitudeEl.ElementFromPathExist("RYYIYR/D").ControlClick()
				}
				catch{
				Result := MsgBox("Unable to process request, please try again. If the error persists, please restart NoteAssist.","Error",5+16+8192)
				if Result = "Retry"
				mancall()
				Exit
				}
				PhnSrch()
			}
			else	
				MsgBox "Latitude not found","Error",48+262144
		}
		
		else if settings["esls"] = 1{
			If WinExist("Latitude - Account Work Form"){
				WinActivate "Latitude - Account Work Form"
				SendInput "!iPPPP{ENTER}"
				sleep 200
				try{
				LatitudeEl.ElementFromPathExist("RYYIYR/D").ControlClick()
				}
				catch{
				Result := MsgBox("Unable to process request, please try again. If the error persists, please restart NoteAssist.","Error",5+16+8192)
				if Result = "Retry"
				mancall()
				Exit
				}
				PhnSrch()
			}
			else	
				MsgBox "Latitude not found","Error",48+262144
		}

}

/*************
*  BUTTON 7  *
*************/
EPAalr := 0
bclick7(Button,Info){
scbval := 0

For bn, label in noteBSCMap{
noteGui["sb" bn].Visible := 0		
}
noteGui["gb1"].Visible := 0
noteGui["gb2"].Visible := 0

If noteGui["FCTitle"].Visible = 1{
Result := MsgBox("Would you like to close the F/C?","Close F/C?",32+262144 " YesNo")
if Result = "No"
	Return
else
CloseFC()
}

If noteGui["slepaTitle"].Visible = 1{
Result := MsgBox("Would you like to close the SLEPA?","Close SLEPA?",32+262144 " YesNo")
if Result = "No"
Return
Else
CloseSLEPA()
}

If noteGui["CalcTitle"].Visible = 1{
Result := MsgBox("Would you like to close the SIF Calculator","Close SIF?",32+262144 " YesNo")
if Result = "No"
Return
else
CalcClose()
}

If noteGui["atytitle"].Visible = 1{
Result := MsgBox("Would you like to close the Attorney note","Close aty?",32+262144 " YesNo")
if Result = "No"
Return
else
CloseATY()
}

If EPAalr = 0
EPAalr := 1
Else
EPAalr := 0

If EPAalr = 0{
CloseEPAb(1,1)
}

If EPAalr = 1{
	epahas := 1
	for val, key in epacMap{
		noteGui["epae" val].Visible := epahas
		noteGui["epat" val].Visible := epahas
	}
	; noteGui[""].Visible 		:= epahas
	noteGui["tN"].Visible 		:= epahas
	noteGui["tY"].Visible 		:= epahas
	noteGui["cbOK1"].Visible 	:= epahas
	noteGui["cbOK2"].Visible 	:= epahas
	noteGui["epadd"].Visible 	:= epahas
	noteGui["cbRec"].Visible 	:= epahas
	noteGui["EPATitle"].Visible := epahas
	noteGui["epab1"].Visible 	:= epahas
	noteGui["epab2"].Visible 	:= epahas
	noteGui["epab3"].Visible 	:= epahas
	noteGui["epab4"].Visible 	:= epahas
	noteGui["epab5"].Visible 	:= epahas
	noteGui["epab6"].Visible 	:= epahas
	
	noteGui["epae7"].Visible 	:= 0
	noteGui["epae9"].Visible 	:= 0
}

noteGui.Hide
noteGui.Show("Autosize") ;Kind of a cheap way to redraw
}

/*************
*  BUTTON 8  *
*************/

FCalr := 0
bclick8(Button,Info){
epahas := 0
scbval := 0
If FCalr = 0
FCalr := 1
Else
FCalr := 0


For bn, label in noteBSCMap{
noteGui["sb" bn].Visible := 0		
}

noteGui["gb1"].Visible := 0
noteGui["gb2"].Visible := 0

If noteGui["EPATitle"].Visible = 1{
	Result := MsgBox("Would you like to close the EPA?","Close EPA?",32+262144 " YesNo")
	if Result = "No"
		Return
	else
	CloseEPA()
}

If noteGui["slepaTitle"].Visible = 1{
Result := MsgBox("Would you like to close the SLEPA?","Close SLEPA?",32+262144 " YesNo")
if Result = "No"
Return
Else
CloseSLEPA()
}

If noteGui["CalcTitle"].Visible = 1{
Result := MsgBox("Would you like to close the SIF Calculator","Close SIF?",32+262144 " YesNo")
if Result = "No"
Return
else
CalcClose()
}

If noteGui["atytitle"].Visible = 1{
Result := MsgBox("Would you like to close the Attorney note","Close aty?",32+262144 " YesNo")
if Result = "No"
Return
else
CloseATY()
}

If FCalr = 0{
CloseFCb(1,1)
}

If FCalr = 1{
	For key, val in fccMap{
	; noteGui["fce" key].Visible 	:= 1
	; noteGui["fce" key].Value 	:= ""
	noteGui["fct" key].Visible 	:= 1
	}
	
	For key, val in fccbMap{
	; noteGui["fce" key].Visible 	:= 1
	; noteGui["fce" key].Value 	:= ""
	noteGui[key].Visible 	:= 1	
	noteGui["t" key].Visible 	:= 1
	}
	
	noteGui["FCTitle"].Visible 	:= 1
	noteGui["fcb1"].Visible 	:= 1
	noteGui["fcb2"].Visible 	:= 1	
	noteGui["fcb3"].Visible 	:= 1
	noteGui["fceMS"].Visible 	:= 1
	; noteGui["fceRO"].Visible 	:= 1
	noteGui["fceEM"].Visible 	:= 1
	noteGui["fceBN"].Visible 	:= 1	

	noteGui.Hide
	noteGui.Show("Autosize")
}

}

/*************
*  BUTTON 9  *
*************/

bclick9(Button,Info){

	If WinExist("Latitude - Account Work Form"){
		WinActivate "Latitude - Account Work Form"
		SendInput "!im"
	}
	else	
		MsgBox "Latitude not found","Error",48+262144

}

/**************
*  EPA CLOSE  *
**************/

CloseEPA(){
	epahas := 0
	for val, key in epacMap{
		noteGui["epae" val].Visible := epahas
		noteGui["epat" val].Visible := epahas
		noteGui["epae" val].Value 	:= ""
	}
	;noteGui["epaSig"].Value := ""
	;noteGui[""].Visible 		:= epahas
	noteGui["tN"].Visible 		:= epahas
	noteGui["tY"].Visible 		:= epahas
	noteGui["cbOK1"].Visible 	:= epahas
	noteGui["cbOK2"].Visible 	:= epahas
	noteGui["epadd"].Visible 	:= epahas
	noteGui["cbRec"].Visible 	:= epahas
	noteGui["EPATitle"].Visible := epahas
	noteGui["epab1"].Visible 	:= epahas
	noteGui["epab2"].Visible 	:= epahas
	noteGui["epab3"].Visible 	:= epahas
	noteGui["epab4"].Visible 	:= epahas
	noteGui["epab5"].Visible 	:= epahas
	noteGui["epab6"].Visible 	:= epahas
	noteGui["cbOK1"].Text 		:= ""
	noteGui["cbOK2"].Text 		:= ""
	noteGui["cbRec"].Text 		:= ""
	noteGui["epadd"].Text		:= ""
	noteGui["epadd"].Enabled 	:= epahas
	noteGui.Hide
	noteGui.Show("Autosize") ;Kind of a cheap way to redraw
	EPAalr := 0
}

CloseEPAb(Button, Info){
	EPAalr := 0
	epahas := 0
	Result := MsgBox("Would you like to close the EPA?","Close?",32+262144 " YesNo")
	if Result = "No"
		Return
	else
	for val, key in epacMap{
		noteGui["epae" val].Visible := epahas
		noteGui["epat" val].Visible := epahas
		noteGui["epae" val].Value 	:= ""
	}
	;noteGui["epaSig"].Value := ""
	;noteGui[""].Visible 		:= epahas
	noteGui["tN"].Visible 		:= epahas
	noteGui["tY"].Visible 		:= epahas
	noteGui["cbOK1"].Visible 	:= epahas
	noteGui["cbOK2"].Visible 	:= epahas
	noteGui["epadd"].Visible 	:= epahas
	noteGui["cbRec"].Visible 	:= epahas
	noteGui["EPATitle"].Visible := epahas
	noteGui["epab1"].Visible 	:= epahas
	noteGui["epab2"].Visible 	:= epahas
	noteGui["epab3"].Visible 	:= epahas
	noteGui["epab4"].Visible 	:= epahas
	noteGui["epab5"].Visible 	:= epahas
	noteGui["epab6"].Visible 	:= epahas
	
	noteGui["cbOK1"].Text 		:= ""
	noteGui["cbOK2"].Text 		:= ""
	noteGui["cbRec"].Text 		:= ""
	noteGui["epadd"].Text		:= ""
	noteGui["epadd"].Enabled 	:= epahas
	
	noteGui.Hide
	noteGui.Show("Autosize") ;Kind of a cheap way to redraw
}

/*****************
*  DiclosureEPA  *
*****************/

DiscV := 0

DiclosureEPA(Button, Info){
epahas := 0
If DiscV >= 1{
noteGui["epaDisc"].GetPos(&sigx,&sigy,&sigw,&sigh)


		for val, key in epacMap{
			noteGui["epae" val].Visible := epahas
			noteGui["epat" val].Visible := epahas
		}
		;noteGui[""].Visible 		:= epahas
		noteGui["tN"].Visible 		:= epahas
		noteGui["tY"].Visible 		:= epahas
		noteGui["cbOK1"].Visible 	:= epahas
		noteGui["cbOK2"].Visible 	:= epahas
		noteGui["epadd"].Visible 	:= epahas
		noteGui["cbRec"].Visible 	:= epahas
		noteGui["EPATitle"].Visible := epahas
		noteGui["epab1"].Visible 	:= epahas
		noteGui["epab2"].Visible 	:= epahas
		noteGui["epab3"].Visible 	:= epahas
		noteGui["epab4"].Visible 	:= epahas
		noteGui["epab5"].Visible 	:= epahas
		noteGui["epab6"].Visible 	:= epahas
		
		noteGui.Hide
		noteGui.Show("Autosize") ;Kind of a cheap way to redraw

noteGui["discb2"].Visible 	:= 1
noteGui["discb1"].Visible 	:= 1
noteGui["epaSig"].Visible 	:= 1
noteGui["epaSigT"].Visible 	:= 1
noteGui["epaDisc"].Visible 	:= 1
noteGui["epat9"].Visible 	:= 1
noteGui["epat9"].Move(sigx+Margin*4,sigy+sigh+margin*4,,)

noteGui["cbOK1"].Visible 	:= 1
noteGui["cbOK1"].Move(sigx+nttw*1.2+Margin*5,sigy+sigh+margin*4+(Margin/10),,)

noteGui["tY"].Visible 		:= 1
noteGui["tY"].Move(sigx+nttw*1.2+Margin*7,sigy+sigh+margin*4,,)

noteGui["cbOK2"].Visible 	:= 1
noteGui["cbOK2"].Move(sigx+nttw*2+Margin*6,sigy+sigh+margin*4+(Margin/10),,)

noteGui["tN"].Visible 		:= 1
noteGui["tN"].Move(sigx+nttw*2.25+Margin*6,sigy+sigh+margin*4,,)

noteGui.Hide
noteGui.Show("Autosize")
}
Else{

		for val, key in epacMap{
			noteGui["epae" val].Visible := epahas
			noteGui["epat" val].Visible := epahas
		}
		;noteGui[""].Visible 		:= epahas
		noteGui["tN"].Visible 		:= epahas
		noteGui["tY"].Visible 		:= epahas
		noteGui["cbOK1"].Visible 	:= epahas
		noteGui["cbOK2"].Visible 	:= epahas
		noteGui["epadd"].Visible 	:= epahas
		noteGui["cbRec"].Visible 	:= epahas
		noteGui["EPATitle"].Visible := epahas
		noteGui["epab1"].Visible 	:= epahas
		noteGui["epab2"].Visible 	:= epahas
		noteGui["epab3"].Visible 	:= epahas
		noteGui["epab4"].Visible 	:= epahas
		noteGui["epab5"].Visible 	:= epahas
		noteGui["epab6"].Visible 	:= epahas
		
		noteGui.Hide
		noteGui.Show("Autosize") ;Kind of a cheap way to redraw


noteGui["EPATitle"].GetPos(&etx,&ety,&etw,&eth)
noteGui.Add("Edit","x" etx " y" ety " w" etw " h" eth*13 " ReadOnly vepaDisc",)

noteGui.Add("Text","x" etx+Margin " y" ety+eth*13+Margin " w" nttw " h" nteh " 0X200 Border Center vepaSigT","EPA:")
noteGui.Add("Edit","xp" nttw " y" ety+eth*13+Margin " w" ntew " h" nteh " Uppercase vepaSig",)

noteGui["b9"].GetPos(&rx,&ry,&rw,&b9h)
noteGui.Add("Button","x" rx+etw/2 " y" ry " w" btnw " h" btnh " vdiscb1","Close",).OnEvent("Click", DiscClose)
noteGui.Add("Button","x" rx+etw/1.17 " y" ry " w" btnw " h" btnh " vdiscb2","Save").OnEvent("Click", DiscSave)
; CloseEPA(1,1)

noteGui["epaDisc"].GetPos(&sigx,&sigy,&sigw,&sigh)

noteGui["epat9"].Visible 	:= 1
noteGui["epat9"].Move(sigx+Margin*4,sigy+sigh+margin*4,,)

noteGui["cbOK1"].Visible 	:= 1
noteGui["cbOK1"].Move(sigx+nttw*1.2+Margin*5,sigy+sigh+margin*4+(Margin/10),,)

noteGui["tY"].Visible 		:= 1
noteGui["tY"].Move(sigx+nttw*1.2+Margin*7,sigy+sigh+margin*4,,)

noteGui["cbOK2"].Visible 	:= 1
noteGui["cbOK2"].Move(sigx+nttw*2+Margin*6,sigy+sigh+margin*4+(Margin/10),,)

noteGui["tN"].Visible 		:= 1
noteGui["tN"].Move(sigx+nttw*2.25+Margin*6,sigy+sigh+margin*4,,)

noteGui.Hide
noteGui.Show("Autosize")
DiscV := DiscV+1
	}

noteGui["cbOK1"].Value := ""
noteGui["cbOK2"].Value := ""
noteGui["epaSig"].Value := ""	
	
SingleDisc := "[" noteGui["epae1"].Value "] by providing your payment details and verbal authorization today [" tday "] you are authorizing Professional Finance Company to create a [SINGLE] electronic payment transaction in the amount of [" noteGui["epae6"].Value "].`n`nThis payment will be processed electronically from your [" noteGui["epae5"].Value "] account ending in [" noteGui["epae4"].Value "].`n`nTo authorize Professional Finance Company to proceed with this electronic payment agreement, we need to capture your voice authorization. At this time, please state your full name and the last four digits of you social security number.`n`nIf there are any issues with the payment, are you ok if we give you a call back?`n`nDo you have any questions about this process?"

RecDisc    := "[" noteGui["epae1"].Value "] by providing your payment details and verbal authorization today [" tday "] you are authorizing Professional Finance Company to create a [RECURRING] electronic payment transaction in the amount of [" noteGui["epae6"].Value "].`n`nThis payment will be processed electronically [" noteGui["epadd"].Text "] from your [" noteGui["epae5"].Value "] account ending in [" noteGui["epae4"].Value "].`n`nTo authorize Professional Finance Company to proceed with this electronic payment agreement, we need to capture your voice authorization. At this time, please state your full name and the last four digits of you social security number.`n`nIf there are any issues with the payment, are you ok if we give you a call back?`n`nWe will send you a written document with the details of this electronic payment agreement.`n`nWe will also email you a notice 5 to 7 days prior to the payment date. If you need to adjust or cancel this payment, you will need to call our office and speak with a representative a minimum of 24 hours before the payment date. Do you have our phone number?`n`n(855-267-7451)`n`nDo you have any questions about this process?"

If noteGui["cbRec"].Text = checkmark{
noteGui["epaDisc"].Value := RecDisc
}Else{
noteGui["epaDisc"].Value := SingleDisc
}
}


/****************
*   DiscClose   *
****************/

DiscClose(Button, Info){
noteGui["discb2"].Visible 	:= 0
noteGui["discb1"].Visible 	:= 0
noteGui["epaSig"].Visible 	:= 0
noteGui["epaSigT"].Visible 	:= 0
noteGui["epaDisc"].Visible 	:= 0
noteGui["epat9"].Move(epat9x, epat9y)
noteGui["cbOK1"].Value := ""
noteGui["cbOK2"].Value := ""
noteGui["epaSig"].Value := ""

if settings["esls"] = 0{
noteGui["t7"].GetPos(&temptx,&tempty,&temptw,&tempth)
noteGui["cbdisp"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epat9"].Move(temptx+temptw+Margin,tempy)
noteGui["cbOK1"].Move(temptx+temptw+nttw+Margin*5,tempy+Margin/4)
noteGui["tY"].Move(temptx+temptw+nttw+Margin*7,tempy+Margin/4)
noteGui["cbOK2"].Move(temptx+temptw+nttw+Margin*17,tempy+Margin/4)
noteGui["tN"].Move(temptx+temptw+nttw+Margin*19,tempy+Margin/4)
;noteGui["epae8"].Value := noteGui["epaSig"].Value
}

Else if settings["esls"] = 1{
noteGui["e5"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epat9"].Move(tempx+tempw+Margin,tempy)
noteGui["cbOK1"].Move(tempx+tempw+nttw+Margin*5,tempy)
noteGui["tY"].Move(tempx+tempw+nttw+Margin*7,tempy)
noteGui["cbOK2"].Move(tempx+tempw+nttw+Margin*17,tempy)
noteGui["tN"].Move(tempx+tempw+nttw+Margin*19,tempy)
;noteGui["epae8"].Value := noteGui["epaSig"].Value
} 

	epahas := 1
	for val, key in epacMap{
		noteGui["epae" val].Visible := epahas
		noteGui["epat" val].Visible := epahas
	}
	; noteGui[""].Visible 		:= epahas
	noteGui["tN"].Visible 		:= epahas
	noteGui["tY"].Visible 		:= epahas
	noteGui["cbOK1"].Visible 	:= epahas
	noteGui["cbOK2"].Visible 	:= epahas
	noteGui["epadd"].Visible 	:= epahas
	noteGui["cbRec"].Visible 	:= epahas
	noteGui["EPATitle"].Visible := epahas
	noteGui["epab1"].Visible 	:= epahas
	noteGui["epab2"].Visible 	:= epahas
	noteGui["epab3"].Visible 	:= epahas
	noteGui["epab4"].Visible 	:= epahas
	noteGui["epab5"].Visible 	:= epahas
	noteGui["epab6"].Visible 	:= epahas
	
	noteGui["epae7"].Visible 	:= 0
	noteGui["epae9"].Visible 	:= 0
}

/***************
*   DiscSave   *
***************/

DiscSave(Button, Info){

noteGui["discb2"].Visible 	:= 0
noteGui["discb1"].Visible 	:= 0
noteGui["epaSig"].Visible 	:= 0
noteGui["epaSigT"].Visible 	:= 0
noteGui["epaDisc"].Visible 	:= 0
noteGui["epat9"].Move(epat9x, epat9y)
if settings["esls"] = 0{
noteGui["t7"].GetPos(&temptx,&tempty,&temptw,&tempth)
noteGui["cbdisp"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epat9"].Move(temptx+temptw+Margin,tempy)
noteGui["cbOK1"].Move(temptx+temptw+nttw+Margin*5,tempy+Margin/4)
noteGui["tY"].Move(temptx+temptw+nttw+Margin*7,tempy+Margin/4)
noteGui["cbOK2"].Move(temptx+temptw+nttw+Margin*17,tempy+Margin/4)
noteGui["tN"].Move(temptx+temptw+nttw+Margin*19,tempy+Margin/4)
noteGui["epae8"].Value := noteGui["epaSig"].Value
;noteGui["epaSig"].Value := ""
}

Else if settings["esls"] = 1{
noteGui["e5"].GetPos(&tempx,&tempy,&tempw,&temph)
noteGui["epat9"].Move(tempx+tempw+Margin,tempy)
noteGui["cbOK1"].Move(tempx+tempw+nttw+Margin*5,tempy)
noteGui["tY"].Move(tempx+tempw+nttw+Margin*7,tempy)
noteGui["cbOK2"].Move(tempx+tempw+nttw+Margin*17,tempy)
noteGui["tN"].Move(tempx+tempw+nttw+Margin*19,tempy)
noteGui["epae8"].Value := noteGui["epaSig"].Value

} 


	epahas := 1
	for val, key in epacMap{
		noteGui["epae" val].Visible := epahas
		noteGui["epat" val].Visible := epahas
	}
	; noteGui[""].Visible 		:= epahas
	noteGui["tN"].Visible 		:= epahas
	noteGui["tY"].Visible 		:= epahas
	noteGui["cbOK1"].Visible 	:= epahas
	noteGui["cbOK2"].Visible 	:= epahas
	noteGui["epadd"].Visible 	:= epahas
	noteGui["cbRec"].Visible 	:= epahas
	noteGui["EPATitle"].Visible := epahas
	noteGui["epab1"].Visible 	:= epahas
	noteGui["epab2"].Visible 	:= epahas
	noteGui["epab3"].Visible 	:= epahas
	noteGui["epab4"].Visible 	:= epahas
	noteGui["epab5"].Visible 	:= epahas
	noteGui["epab6"].Visible 	:= epahas
	
	noteGui["epae7"].Visible 	:= 0
	noteGui["epae9"].Visible 	:= 0
}

/**************
*   PostEPA   *
**************/

PostEPA(Button, Info){
epaStr := ""
epaStrL := "1,4,5,6,8"
; if noteGui["epae" val].Value = ""


Loop parse epaStrL, ","{
if noteGui["epae" A_LoopField].Value = ""
epaStr := epaStr
Else
if A_Index = 2
	epaStr .= epacMap["" A_LoopField] " *" noteGui["epae" A_LoopField].Value delim
Else
if A_Index = 4{
epaStr .= epacMap["6"] " " noteGui["epae6"].Value delim
If noteGui["cbRec"].Text = checkmark
	epaStr .= "Frequency: " noteGui["epadd"].Text delim
}
Else
epaStr .= epacMap["" A_LoopField] " " noteGui["epae" A_LoopField].Value delim
}

If noteGui["cbOK1"].Value = checkmark
epaStr .= "OK TO CB"
Else If noteGui["cbOK2"].Value = checkmark
epaStr .= "NOT OK TO CB"
	
noteGui["e7"].Focus()
SendInput "^{End}"
EditPaste epaStr, noteGui["e7"]

If noteGui["cbRec"].Text = checkmark{
slepapost := MsgBox("You have selected a recurring payment, would you like to submit a SLEPA?","Post SLEPA?",32+262144 " YesNo")
If slepapost = "No"
	Return
Else
	Result := MsgBox("Would you like to close the EPA?","Close?",32+262144 " YesNo")
	if Result = "No"
		Return
	else{
	CloseEPA()
	sbclick005(1,1)
  }
 }
}

/**************
*   EPAIMP2   *
**************/

assettab(asset := "",title := ""){
LatitudeE3 := UIA.ElementFromHandle("Debtor Details")
LatitudeE3.ElementFromPath("Y").SetFocus()
SendInput "+{HOME}" "+{Right}{Right}{Right}{Right}{Right}"
LatitudeE3 := UIA.ElementFromHandle("Debtor Details")
LatitudeE3.WaitElement({Type: "CheckBox", Name: "Value Verified"})

LatitudeE3.ElementFromPath("YY/YY").SetFocus()
SendInput "+{END}"
LatitudeE3.ElementFromPath("YY/YYYR4").Value := asset " - " tday
LatitudeE3.ElementFromPath("YY/YYYR3").Value := title
SendInput "!v"
}

EPAimp2(Button, Info){
If WinExist("Latitude - Account Work Form"){
imp2r := MsgBox("Would you like to save the bank info?","Bank",32+262144 " YesNo")
If imp2r = "No"
	Return
Else if noteGui["epae5"].Value = ""
	MsgBox "Missing BANK NAME. Please enter the bank name and try again.","Error",48+262144
Else if WinExist("Debtor Details"){
		WinShow "Debtor Details"
		WinActivate "Debtor Details"
		WinWaitActive "Debtor Details",,3
		assettab(noteGui["epae5"].value,"BANK - Bank Account")
		sleep 300
		SendInput "!o"
	}Else{
		goodbyepops()
		WinActivate "Latitude - Account Work Form"
		SendInput "!rv"
		WinWait "Debtor Details",,3
		WinActivate "Debtor Details"
		WinWaitActive "Debtor Details",,3
		assettab(noteGui["epae5"].value,"BANK - Bank Account")
		sleep 300
		SendInput "!o"
	}
		sleep 100
		MsgBox "Bank information saved successfully!","Success",32+262144
}
else
MsgBox
}


/**************
*   PostF/C   *
**************/

PostFC(Button, Info){

	; Global fccbMap := Map(
	; "cb1ms",	"Single"		,
	; "cb2ms",	"Married"		,
	; "cb3ro",	"Rent"			,
	; "cb4ro",	"Own"			,
	; "cb5em",	"Unemployed"	,
	; "cb6em",	"Employed"		,
	; "cb7bn",	"Refused"		,
	; "cb8bn",	"Consent"		, 
	; )

fcStr := delim "F/C: "

If notegui["cb1ms"].Text = checkmark
fcStr .= fccbMap["cb1ms"] delim

If notegui["cb2ms"].Text = checkmark
fcStr .= fccbMap["cb2ms"] ": " noteGui["fceMS"].Value delim

If notegui["cb3ro"].Text = checkmark
fcStr .= fccbMap["cb3ro"] delim

If notegui["cb4ro"].Text = checkmark
fcStr .= fccbMap["cb4ro"] delim ;": " noteGui["fceRO"].Value delim

If notegui["cb5em"].Text = checkmark
fcStr .= fccbMap["cb5em"] delim

If notegui["cb6em"].Text = checkmark
fcStr .= fccbMap["cb6em"] ": " noteGui["fceEM"].Value delim

If notegui["cb7bn"].Text = checkmark
fcStr .= "BANK NAME: " fccbMap["cb7bn"] delim

If notegui["cb8bn"].Text = checkmark
fcStr .= "BANK NAME: " noteGui["fceBN"].Value delim

noteGui["e7"].Focus()
SendInput "^{End}"
EditPaste fcStr, noteGui["e7"]
}

/***************
*   CloseF/C   *
***************/

CloseFC(){

	For key, val in fccMap{
	; noteGui["fce" key].Visible := 0
	; noteGui["fce" key].Value := ""
	noteGui["fct" key].Visible := 0
	}
	For key, val in fccbMap{
	; noteGui["fce" key].Visible := 0
	; noteGui["fce" key].Value := ""
	noteGui[key].Visible := 0
	noteGui["t" key].Visible := 0
	}
	
	noteGui["fceMS"].Value 	:= ""
	; noteGui["fceRO"].Value 	:= ""
	noteGui["fceEM"].Value 	:= ""
	noteGui["fceBN"].Value 	:= ""
	
	notegui["cb1ms"].Value 	:= ""	
	notegui["cb2ms"].Value 	:= ""
	notegui["cb3ro"].Value 	:= ""
	notegui["cb4ro"].Value 	:= ""
	notegui["cb5em"].Value 	:= ""
	notegui["cb6em"].Value 	:= ""
	notegui["cb7bn"].Value 	:= ""
	notegui["cb8bn"].Value 	:= ""
	
	
	noteGui["FCTitle"].Visible 	:= 0
	noteGui["fcb1"].Visible 	:= 0
	noteGui["fcb2"].Visible 	:= 0
	noteGui["fcb3"].Visible 	:= 0
	noteGui["fceMS"].Visible 	:= 0
	; noteGui["fceRO"].Visible 	:= 0
	noteGui["fceEM"].Visible 	:= 0
	noteGui["fceBN"].Visible 	:= 0
	
	noteGui["fceMS"].Enabled 	:= 0
	; noteGui["fceRO"].Enabled 	:= 0
	noteGui["fceEM"].Enabled 	:= 0
	noteGui["fceBN"].Enabled 	:= 0	

	noteGui.Hide
	noteGui.Show("Autosize")
	FCalr := 0
}

CloseFCb(Button, Info){

	Result := MsgBox("Would you like to close the F/C?","Close?",32+262144 " YesNo")
	if Result = "No"
		Return
	else
	For key, val in fccMap{
	; noteGui["fce" key].Visible := 0
	; noteGui["fce" key].Value := ""
	noteGui["fct" key].Visible := 0
	}
	For key, val in fccbMap{
	; noteGui["fce" key].Visible := 0
	; noteGui["fce" key].Value := ""
	noteGui[key].Visible := 0
	noteGui["t" key].Visible := 0
	}
	
	noteGui["fceMS"].Value 	:= ""
	; noteGui["fceRO"].Value 	:= ""
	noteGui["fceEM"].Value 	:= ""
	noteGui["fceBN"].Value 	:= ""
	
	notegui["cb1ms"].Value 	:= ""	
	notegui["cb2ms"].Value 	:= ""
	notegui["cb3ro"].Value 	:= ""
	notegui["cb4ro"].Value 	:= ""
	notegui["cb5em"].Value 	:= ""
	notegui["cb6em"].Value 	:= ""
	notegui["cb7bn"].Value 	:= ""
	notegui["cb8bn"].Value 	:= ""
	
	
	noteGui["FCTitle"].Visible 	:= 0
	noteGui["fcb1"].Visible 	:= 0
	noteGui["fcb2"].Visible 	:= 0
	noteGui["fcb3"].Visible 	:= 0
	noteGui["fceMS"].Visible 	:= 0
	; noteGui["fceRO"].Visible 	:= 0
	noteGui["fceEM"].Visible 	:= 0
	noteGui["fceBN"].Visible 	:= 0
	
	noteGui["fceMS"].Enabled 	:= 0
	; noteGui["fceRO"].Enabled 	:= 0
	noteGui["fceEM"].Enabled 	:= 0
	noteGui["fceBN"].Enabled 	:= 0	

	noteGui.Hide
	noteGui.Show("Autosize")
	FCalr := 0
}

/****************
*   FILL F/C    *
****************/

FillFC(Button, Info){
If WinExist("Latitude - Account Work Form"){
fcpostr := MsgBox("Would you like to autofill the Full and Complete","F/C Fill?",32+262144 " YesNo")
If fcpostr = "No"
	Return

Else if WinExist("Debtor Details"){
WinShow "Debtor Details"
WinActivate "Debtor Details"
WinWaitActive "Debtor Details",,3
fillfcfunc()
MsgBox "Done","Done",64+262144
}

Else if !WinExist("Debtor Details"){
goodbyepops()
WinActivate "Latitude - Account Work Form"
WinWaitActive "Latitude - Account Work Form",,3
SendInput "!rv"
WinWait "Debtor Details",,3
WinActivate "Debtor Details"
WinWaitActive "Debtor Details"
fillfcfunc()
MsgBox "Done","Done",64+262144
}
}
else	
MsgBox "Latitude not found","Error",48+262144
}

fillfcfunc(){

If notegui["cb2ms"].Text = checkmark
marsin := noteGui["fceMS"].Value
else
marsin := ""

If notegui["cb4ro"].Text = checkmark
renown := "OWN"
else
renown := ""

If notegui["cb6em"].Text = checkmark
emplfc := noteGui["fceEM"].Value
else
emplfc := ""

If notegui["cb8bn"].Text = checkmark
bankfc := noteGui["fceBN"].Value
else
bankfc := ""

SendInput "!d" "+{Home}" "!d" "+{Home}"

sleep 100

/**********/
/*Employed*/
/**********/

try{
if emplfc = ""{ ;this is the only time we check to see if the first box is checked - if it is we are going to delete the info already there
If notegui["cb6em"].Text = checkmark{
return
}
Else{
LatitudeE3 := UIA.ElementFromHandle("Debtor Details ahk_exe Latitude.exe")
LatitudeE3.ElementFromPath("YY/Y4").Value := emplfc
;msgbox "Save Button: " LatitudeE3.ElementFromPath("YY0").IsEnabled
SendInput "!v"
while LatitudeE3.ElementFromPath("YY0").IsEnabled = 1{
sleep 10
if LatitudeE3.ElementFromPath("YY0").IsEnabled = 0{
break
   }
  }
 }
}
Else{
SendInput "!d"
LatitudeE3 := UIA.ElementFromHandle("Debtor Details ahk_exe Latitude.exe")
LatitudeE3.ElementFromPath("YY/Y4").Value := emplfc " " tday
;msgbox "Save Button: " LatitudeE3.ElementFromPath("YY0").IsEnabled
SendInput "!v"
while LatitudeE3.ElementFromPath("YY0").IsEnabled = 1{
sleep 10
if LatitudeE3.ElementFromPath("YY0").IsEnabled = 0{
break
  }
 }
}
sleep 100
}

  /*********/
 /*MARSING*/
/*********/

try{
if marsin = ""
marsin := marsin
Else{
SendInput "!s"
sleep 50
LatitudeE3 := UIA.ElementFromHandle("Debtor Details ahk_exe Latitude.exe")
if LatitudeE3.ElementFromPath("1,2,3").Type != 50004{
while LatitudeE3.ElementFromPath("1,2,3").Type != 50004{
sleep 10
SendInput "!s"
if LatitudeE3.ElementFromPath("1,2,3").Type = 50004{
LatitudeE3.ElementFromPath("1,2,3").Value := marsin " " tday
break
  }
 }
}
Else{
LatitudeE3.ElementFromPath("1,2,3").Value := marsin " " tday
}
SendInput "!v"
while LatitudeE3.ElementFromPath("YY0").IsEnabled = 1{
sleep 10
if LatitudeE3.ElementFromPath("YY0").IsEnabled = 0{
break
  }
 }
}
sleep 100
}

Try{
if renown = ""
renown := renown
Else
assettab(renown, "REAL - Real Property")
}

Try{
if bankfc = ""
bankfc := bankfc
Else
assettab(bankfc,"BANK - Bank Account")
}

} ;end fcfillfunc

/****************
*   PostSLEPA   *
****************/

PostSLEPA(Button, Info){
slepastr := ""
for key, val in slepamap{
sleco := A_index
if noteGui["slepae" key].Value = ""{
MsgBox("You are missing info, please ensure all fields are properly filled out","Error",48+262144)
sleco := sleco - 1
break
}
else
slepastr := slepastr val " " noteGui["slepae" key].Value "; "
}
if sleco < 3
return
else{
result := MsgBox("Would you like to post a SLEPA?","SLEPA",32+262144 " YesNo")
If result = "No"
	Return
Else
slepastr := RTrim(slepastr, "; ")

	If WinExist("Latitude - Account Work Form"){
		WinActivate "Latitude - Account Work Form"
		SendInput "!n"
		sleep 250
		SendInput "!n"
		WinWait "Latitude - New Note"
		If WinExist("Latitude - New Note"){
			WinActivate "Latitude - New Note"
			A_Clipboard := RTrim(slepastr, ";")
			WinActivate "Latitude - New Note"
			SendInput "SL{Down}{Tab}SLEPA{Down}{Tab}"
			SendInput "+{End}{Del}"
			SendInput "^v"
			SendInput "!a"
			slepastr := ""
		}
	}
	else	
	MsgBox "Latitude not found","Error",48+262144	
}


}

/*****************
*   CloseSLEPA   *
*****************/
CloseSLEPA(){
For bn, label in noteBSCMap{
noteGui["sb" bn].Visible := scbval
}
noteGui["gb1"].Visible := scbval
noteGui["gb2"].Visible := scbval
For key, val in slepaMap{
noteGui["slepae" key].Visible 	:= 0
noteGui["slepae" key].Value 	:= ""
noteGui["slepat" key].Visible 	:= 0
}
noteGui["slepaTitle"].Visible 	:= 0
noteGui["slepab1"].Visible 	:= 0
noteGui["slepab2"].Visible  := 0
noteGui["slepab3"].Visible 	:= 0
noteGui["slepab4"].Visible := 0
noteGui.Hide
noteGui.Show("Autosize")
}

CloseSLEPAb(Button, Info){
Result := MsgBox("Would you like to close the SLEPA?","Close SLEPA?",32+262144 " YesNo")
if Result = "No"
Return
Else
For bn, label in noteBSCMap{
noteGui["sb" bn].Visible := scbval
}
noteGui["gb1"].Visible := scbval
noteGui["gb2"].Visible := scbval
For key, val in slepaMap{
noteGui["slepae" key].Visible 	:= 0
noteGui["slepae" key].Value 	:= ""
noteGui["slepat" key].Visible 	:= 0
}
noteGui["slepaTitle"].Visible 	:= 0
noteGui["slepab1"].Visible 	:= 0
noteGui["slepab2"].Visible  := 0
noteGui["slepab3"].Visible 	:= 0
noteGui["slepab4"].Visible := 0
noteGui.Hide
noteGui.Show("Autosize")
}

/******************
*   goodbyepops   *
******************/
goodbyepops(){
try{
if WinExist("Latitude Search")
WinClose("Latitude Search")
if WinExist("Account Warnings")
WinClose("Account Warnings")
if WinExist("All Debtor Notes")
WinClose("All Debtor Notes")
if WinExist("Miscellaneous Extra Data")
WinClose("Miscellaneous Extra Data")
if WinExist("Debtor Contact Analysis")
WinClose("Debtor Contact Analysis")
if WinExist("Debtor Details")
WinClose("Debtor Details")
} ;try
}



/********************
***			      ***
***   SHORTCUTS   ***
***			      ***
********************/

/*************
*    ADDR    *
*************/
sbclick001(Button, Info){
;MsgBox "• ESINV`n• ESPAYER`n• CLOSEINV`n`to If the account was in Late stage`n`tbefore being closed> TRC to Sr`n`tRS`n• DECDESK`n• NBEXCEPTION`n• NEWINV`n• SCRADESK`n`to Consumer currently Active-Duty`nMilitary`n`to No Credit Reporting`n`to No Interest`n• SUSINV`n• LSTRIGGERS`n• BKYDESK`n`to Bankruptcy must be discharged`nor dismissed.`n`to Do not answer question about`naccounts that were included in`nbankruptcy.`n• CMPDESK`n`to w/ SUP approval only"
MsgBox "Physical Address: 5754 W 11th St Ste 100 Greeley CO 80634`n`nMailing Address: PO BOX 1686 Greeley CO, 80632","Address",64+262144
}
/*************
*  PMT Matr  *
*************/
sbclick002(Button, Info){
MsgBox "< $500 - PIF`n`n$501-$2499 - 6 MONTHS`n`n$2500 - 4999 - 12 MONTHS`n`n>$5000 - 24 MONTHS","Payment Matrix",64+262144
}
/*************
*    MM      *
*************/
sbclick003(Button, Info){
MsgBox "MM - My name is [], and I am a debt collector with Professional Finance Company. This is an attempt to collect a debt. Any information obtained will be used for that purpose. The call is recorded. Is it OK to continue?`n`n`nNV6 - If an account is marked NV6 you must tell the consumer: '(1) A payment is not demanded or due; (2) The medical debt will not be reported to any credit reporting agency during the 60-day notification period'","Mini Miranda",64+262144
}
/*************
*   PIFLR    *
*************/
sbclick004(Button, Info){

PIFLRres := MsgBox("Would you like to post a PIFLR","Post PIFLR?",32+262144 " YesNo")
If PIFLRres = "No"
	Return
Else
If noteGui["e4"].Text = ""
MsgBox "Fill in email and try again"
Else
	If WinExist("Latitude - Account Work Form"){
		WinActivate "Latitude - Account Work Form"
		SendInput "!n"
		sleep 250
		SendInput "!n"
		WinWait "Latitude - New Note"
		If WinExist("Latitude - New Note"){
		A_clipboard := noteGui["e4"].Text
		WinActivate "Latitude - New Note"
		SendInput "CO{Down}{Tab}PIFLR{Down}{Tab}"
		SendInput "^v"
		SendInput "!a"
		}
	}
	
	else	
		MsgBox "Latitude not found","Error",48+262144	
}
/**************
*    SLEPA    *
**************/
sbclick005(Button, Info){
scbval := 0
For bn, label in noteBSCMap{
noteGui["sb" bn].Visible := 0		
}

noteGui["gb1"].Visible := 0
noteGui["gb2"].Visible := 0

If noteGui["SLEPATitle"].Visible = 1
return
Else
For key, val in slepaMap{
noteGui["slepae" key].Visible 	:= 1
noteGui["slepae" key].Value 	:= ""
noteGui["slepat" key].Visible 	:= 1
}
noteGui["slepaTitle"].Visible 	:= 1
noteGui["slepab1"].Visible 	:= 1
noteGui["slepab2"].Visible  := 1
noteGui["slepab3"].Visible 	:= 1
noteGui["slepab4"].Visible  := 1

noteGui.Hide
noteGui.Show("Autosize")

}
/*************
*  Dispute   *
*************/
sbclick006(Button, Info){
if WinExist("Latitude - Account Work Form"){
	If WinExist("Debtor Details"){
		WinShow "Debtor Details"
		WinActivate "Debtor Details"
		WinWaitActive "Debtor Details"
		SendInput "!d" "+{Home}" "!d" "+{Home}"
		SendInput "!r"
		LatitudeE8 := UIA.ElementFromHandle("Debtor Details ahk_exe Latitude.exe")
		if LatitudeE8.ElementFromPath("YY/Y2q").ToggleState = 0{
		LatitudeE8.ElementFromPath("YY/Y2q").ToggleState := 1
		SendInput "!v"
		while LatitudeE8.ElementFromPath("YY0").IsEnabled = 1{
		sleep 10
		if LatitudeE8.ElementFromPath("YY0").IsEnabled = 0{
		break
		  }
		 }
		}
		Else
		MsgBox "Account is already disputed","Error",48+262144
	  }
	  
	Else{
		WinActivate "Latitude - Account Work Form"
		SendInput "!rv"
		WinWait "Debtor Details",,3
		WinActivate "Debtor Details"
		WinWaitActive "Debtor Details"
		SendInput "!r"
		LatitudeE8 := UIA.ElementFromHandle("Debtor Details ahk_exe Latitude.exe")
		if LatitudeE8.ElementFromPath("YY/Y2q").ToggleState = 0{
		LatitudeE8.ElementFromPath("YY/Y2q").ToggleState := 1
		SendInput "!v"
		while LatitudeE8.ElementFromPath("YY0").IsEnabled = 1{
		sleep 10
		if LatitudeE8.ElementFromPath("YY0").IsEnabled = 0{
		break
		  }
		 }
		}
		Else
		MsgBox "Account is already disputed","Error",48+262144
	  }
	 }
	Else
	MsgBox "Latitude not found","Error",48+262144
}

;mClr := PixelGetColor(150*WinGetDpi("Debtor Details")/144,127*WinGetDpi("Debtor Details")/144) (Reminder that you once checked if the account was disputed by checking the color of a pixel in a check box T_T )
	

/******************
*  Consent Email  *
******************/
sbclick007(Button, Info){

MsgBox "CONSENT@PFCUSA.COM is for the purpose of submitting the following:`n`n1.`tProof of Payment`n2.`tExplanation of benefits`n3.`tMedicaid LOC or Financial Assistance Approval Letters`n4.`tPIF deletion letter requests in special circumstances – `t`tie. Lending situations`n5.`tAuthorizations – POA, CCCS etc.`n6.`tSIF offer requests`n`nThe 'Consent' email is *not* to be used for consumer requests for itemized statements, print outs or disputes. Please do not advise the consumer to send requests for these to consent they may request these in writing via mail or our website using the dispute form.","Consent",64+262144

}
/******************
*  Phone Numbers  *
******************/
sbclick008(Button, Info){

MsgBox "Early Stage: (855) 267-7451, (833) 871-1862, (833)-871-1863`nLate Stage: (855) 267-5572`n`n**Credit Bureaus**`nEquifax: (800) 685-1111`nExperian: (888) 397-3742`nTransUnion: (800)-888-4213`n`nFax Number: (866) 644-3420`n`nOutbound Message Script:`n'Hello, this is a message from Professional Finance Company, a debt collector. Please give us a call back at 855-267-7451.'","Phone Numbers",64+262144

}
/*****************
*  Restrictions  *
*****************/
sbclick009(Button, Info){

MsgBox "*No Message States:*`n-Alabama`t-Buffalo, NY`n-Florida`t`t-Georgia`n-New Hampshire`t-New York City, NY`n`n*No Calls At Work:*`n-North Carolina`n-New Hampshire - 1 Call at work per month`n-Oregon - No calls unless unable to reach at home`n-Washington - 1 Call per week only at work`n`n*No Spousal Contact:*(Without Consent)`n-Arizona`t`t-Connecticut`n-Hawaii`t`t-Georgia`n`n*Time-Barred Debt:* (Out of Statute)`nCannot be collected in:`nMississippi and Wisconsin`n`n*Time-Barred Debt disclosures are required in:*`n-California`t-Connecticut`n-Massachusetts`t-North Carolina`n-New Mexico`t-New York City`n`nRefusal to Pay = CEASE COMM`nAny reason not to pay = DISPUTE","Restrictions",64+262144

}

/*************
*     NA     *
*************/
sbclick010(Button, Info){
if WinExist("Latitude - Account Work Form"){
try{
Result := MsgBox("Post NA?","NA?",32+262144 " YesNo")
if Result = "No"
	Return
else
if notegui["cbbad"].Value = checkmark
Result := MsgBox("The Bad# box is checked. This will mark the number as bad. Are you sure you would like to continue?","Bad Number",262144 " YesNo")
if Result = "No"
	Return
else
obcall()
if LatitudeEl.ElementFromPathExist("RYYIYR3/K"){
if LatitudeEl.ElementFromPath("RYYIYR3/K").Type = 50020{
if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "NA - NO ANSWER"
LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "NA - NO ANSWER"
if LatitudeEl.ElementFromPath("RYYIYR4").Value != "NA - NO ANSWER"
LatitudeEl.ElementFromPath("RYYIYR4").Value := "NA - NO ANSWER"

LatitudeEl.ElementFromPath("RYYIYR0/").ControlClick()
;sleep 150
WinWait "Note Entry/Update Confirmation"
WinClose "Note Entry/Update Confirmation"
   }
  }
 } ;TRY
catch as e{
MsgBox "There was an error processing your request. If the issue persists, please restart NoteAssist.","Error 0x" e.line,16+262144
}
} ;winexist
Else
MsgBox "Latitude not found","Error",48+262144
}

/*************
*    PUHU    *
*************/

sbclick011(Button, Info){
if WinExist("Latitude - Account Work Form"){
Try{
Result := MsgBox("Post PUHU?","PUHU?",32+262144 " YesNo")
if Result = "No"
	Return
else
; if notegui["cbbad"].Value = checkmark
; Result := MsgBox("The Bad# box is checked. This will mark the number as bad. Are you sure you would like to continue?","Bad Number",16+262144 " YesNo")
; if Result = "No"
	; Return
; else
obcall()
if LatitudeEl.ElementFromPathExist("RYYIYR3/K"){
if LatitudeEl.ElementFromPath("RYYIYR3/K").Type = 50020{
if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "PUHU - PICK UP HUNG UP"
LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "PUHU - PICK UP HUNG UP"
if LatitudeEl.ElementFromPath("RYYIYR4").Value != "PUHU - PICK UP HANG UP"
LatitudeEl.ElementFromPath("RYYIYR4").Value := "PUHU - PICK UP HANG UP"
LatitudeEl.ElementFromPath("RYYIYR0/").ControlClick()
WinWait "Note Entry/Update Confirmation"
WinClose "Note Entry/Update Confirmation"
   }
  }
 } ;try
catch as e{
MsgBox "There was an error processing your request. If the issue persists, please restart NoteAssist.","Error 0x" e.line,16+262144
}
} ;winexist
Else
MsgBox "Latitude not found","Error",48+262144
}

/*************
*    RTV     *
*************/
sbclick012(Button, Info){
if WinExist("Latitude - Account Work Form"){
rtvname := notegui["e1"].Value
rtvnamea := StrSplit(rtvname, " ")

if rtvname = ""{
MsgBox "Please fill out NAME and try again","Error: Missing Name",48+262144
return
}
else{
Try{
OnMessage(0x44, rtvibobmsg)
global rtvMsgBoxTitle := "Inbound/Outbound"
rtvibobresult := MsgBox("Please select inbound or outbound:", rtvMsgBoxTitle, 64+262144 " YesNoCancel")


rtvibobmsg(wParam, lParam, msg, hwnd){
    global rtvMsgBoxTitle
    DetectHiddenWindows true
    rtvmsgBoxHwnd := WinExist(rtvMsgBoxTitle)
    if (rtvmsgBoxHwnd) {
      ControlSetText "Inbound", "Button1", rtvmsgBoxHwnd 	;yes
      ControlSetText "Outbound", "Button2", rtvmsgBoxHwnd 	;no
    }
    DetectHiddenWindows false
}

if rtvibobresult = "Cancel"{
return
}
else{
if rtvibobresult = "Yes"{ ;inbound
inbcall()
}
else if rtvibobresult = "No"{ ;outbound
obcall()
 }
}

if noteGui["e2"].Value = ""
addr := ""
else 
addr := noteGui["t2"].Text " " noteGui["e2"].Value delim

if noteGui["cb1DOB"].Text = ""
dob := ""
else 
dob := "DOB" delim

if noteGui["cb2SSN"].Text = ""
ssn := ""
else 
ssn := "SSN" delim

If noteGui["cb3MM"].Text = ""{
if rtvibobresult = "Yes"{
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "RTV - REFUSED TO VERIFY"
LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "RTV - REFUSED TO VERIFY"
LatitudeEl.ElementFromPath("RYYIYR4").Value := RTVNAMEA[1] delim "RTV" delim addr dob ssn notegui["e7"].Value 
}

if rtvibobresult = "No"{
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "RTV - REFUSED TO VERIFY"
LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "RTV - REFUSED TO VERIFY"
LatitudeEl.ElementFromPath("RYYIYR4").Value := RTVNAMEA[1] delim "RTV" delim notegui["e7"].Value 
}
}
else{
if rtvibobresult = "Yes"{
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "RTV - REFUSED TO VERIFY"
LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "RTV - REFUSED TO VERIFY"
LatitudeEl.ElementFromPath("RYYIYR4").Value := RTVNAMEA[1] delim addr dob ssn noteCBm["cb3MM"] delim "RTV" delim notegui["e7"].Value 
}

if rtvibobresult = "No"{
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "RTV - REFUSED TO VERIFY"
LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "RTV - REFUSED TO VERIFY"
LatitudeEl.ElementFromPath("RYYIYR4").Value := RTVNAMEA[1] delim noteCBm["cb3MM"] delim "RTV" delim notegui["e7"].Value 
 }
}
SendInput "^{END}"
   } ;try
   catch as e{
   MsgBox "There was an error processing your request. If the issue persists, please restart NoteAssist.","Error 0x" e.line,16+262144
   }
  } ;else
 } ;if lat exist
Else
MsgBox "Latitude not found","Error",48+262144
}

/*****************
*  Search Phone  *
*****************/
sbclick013(Button, Info){
		notegui.opt("-AlwaysOnTop")
		
		if WinExist("Latitude - Account Work Form"){
		try{
		agntph := ""
		if WinExist("ahk_exe Agent Desktop Native.exe"){
		try{
		;MsgBox "try1 phone"
		AgentEl := UIA.ElementFromHandle("ahk_exe Agent Desktop Native.exe")
		if AgentEl.ElementFromPathExist("VKr"){
		agntph := AgentEl.ElementFromPath("VKr").Name
		agntph := RegExReplace(agntph, "\D", "")
		 }
		}
		
		if agntph = ""{
		try{
		AgentEl := UIA.ElementFromHandle("ahk_exe Agent Desktop Native.exe")
		if AgentEl.ElementFromPathExist("VKs"){
		agntph := AgentEl.ElementFromPath("VKs").Name
		agntph := RegExReplace(agntph, "\D", "")
		  }
		 }
		}
		
		if agntph = ""{		
		try{
		;MsgBox "try1 phone"
		AgentEl := UIA.ElementFromHandle("Agent Panel - Main window")
		if AgentEl.ElementFromPathExist("VKr"){
		agntph := AgentEl.ElementFromPath("VKr").Name
		agntph := RegExReplace(agntph, "\D", "")
		  }
		 }
		}
		
		if agntph = ""{
		try{
		AgentEl := UIA.ElementFromHandle("Agent Panel - Main window")
		if AgentEl.ElementFromPathExist("VKs"){
		agntph := AgentEl.ElementFromPath("VKs").Name
		agntph := RegExReplace(agntph, "\D", "")
		  }
		 }
		}
		
		;MsgBox agntph
		if agntph != ""{
		agntph := RegExReplace(agntph, "\D", "")
		  }
		 Else
		 MsgBox "Could not get number from LiveVox","Error",48+262144
		 }
		Else
		MsgBox "Could not find LiveVox","Error",48+262144
		
		agntph := RegExReplace(agntph, "\D", "")
		if agntph = ""
		msgbox "Unable to retrive phone number from Livevox. Please enter a phone number and try again.","Error",48+262144
		Else{
		if WinExist("Latitude - Account Work Form"){
		WinActivate "Latitude - Account Work Form"
		SendInput "^f"
		WinWait "Latitude Search",,3
		WinActivate "Latitude Search"
		WinWaitActive "Latitude Search"
		LatitudeE5 := UIA.ElementFromHandle("Latitude Search")
		if LatitudeE5.ElementFromPathExist("Y0/").Name = "Back"{
		LatitudeE5.ElementFromPathExist("Y0/").ControlClick()
		LatitudeE5.ElementFromPathExist("Y0").ControlClick()
		sleep 100
		SendInput "!y"
		 }
		}
		Else{
		notegui.opt("-AlwaysOnTop")
		LatitudeE5.ElementFromPath("Y0").ControlClick()
		WinWait "Latitude Search",,3
		WinActivate "Latitude Search"
		WinWaitActive "Latitude Search"
		SendInput "!y"
		}
		LatitudeE5.ElementFromPathExist("IYYYxY3").Value := "Phone History"
		agntph := RegExReplace(agntph, "\D", "")
		LatitudeE5.ElementFromPathExist("IYYYxY4").Value := agntph
		LatitudeE5.ElementFromPathExist("Y0/").ControlClick()
		  }
		 } ;try
		catch as e{
		MsgBox "There was an error processing your request. If the issue persists, please restart NoteAssist.","Error 0x" e.line,16+262144
		}
		} ;winex
		Else
		MsgBox "Latitude not found","Error",48+262144 
		if settings["aaot"] = 1
		notegui.opt("AlwaysOnTop")
		else
		notegui.opt("-AlwaysOnTop")
}

/*************
*   MCAID    *
*************/
sbclick014(Button, Info){
Result := MsgBox("Would you like to post MCAID","Post MCAID?",32+262144 " YesNo")
If Result = "No"
	Return
Else
	If WinExist("Latitude - Account Work Form"){
		WinActivate "Latitude - Account Work Form"
		SendInput "!n"
		sleep 100
		SendInput "!n"
	WinWait "Latitude - New Note"
		If WinExist("Latitude - New Note"){
			WinActivate "Latitude - New Note"
			SendInput "CO{Down}{Tab}MCAID{Down}{Tab}"
			SendInput "!a"
		}
	}
	
	else	
		MsgBox "Latitude not found","Error",48+262144	
}

/*************
*   HLD 30   *
*************/
sbclick015(Button, Info){

Result := MsgBox("Would you like to post HLD30","Post HLD30?",32+262144 " YesNo")
If Result = "No"
	Return
Else
	If WinExist("Latitude - Account Work Form"){
		WinActivate "Latitude - Account Work Form"
		SendInput "!n"
		sleep 100
		SendInput "!n"
	WinWait "Latitude - New Note"
		If WinExist("Latitude - New Note"){
			WinActivate "Latitude - New Note"
			SendInput "CO{Down}{Tab}HLD30{Down}{Tab}"
			SendInput "!a"
		}
	}
	else	
	MsgBox "Latitude not found","Error",48+262144	

}

/*************
*    NML     *
*************/
sbclick016(Button, Info){
if WinExist("Latitude - Account Work Form"){
try{

OnMessage(0x44, nmlibobmsg)
global nmlMsgBoxTitle := "Outbound/Manual"
ibobresult := MsgBox("Please select outbound or manual:", nmlMsgBoxTitle, 64+262144 " YesNoCancel")


nmlibobmsg(wParam, lParam, msg, hwnd){
    global nmlMsgBoxTitle
    DetectHiddenWindows true
    msgBoxHwnd := WinExist(nmlMsgBoxTitle)
    if (msgBoxHwnd) {
      ControlSetText "Outbound", "Button1", msgBoxHwnd 	;yes
      ControlSetText "Manual", "Button2", msgBoxHwnd 	;no
    }
    DetectHiddenWindows false
}

if ibobresult = "Cancel"{
return
}
else{
if ibobresult = "Yes"{ ;outbound
obcall()
}
else if ibobresult = "No"{ ;manual
mancall()
 }
}
if LatitudeEl.ElementFromPathExist("RYYIYR3/K"){
if LatitudeEl.ElementFromPath("RYYIYR3/K").Type = 50020{
if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "NML - NO MESSAGE LEFT"
LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "NML - NO MESSAGE LEFT"
if LatitudeEl.ElementFromPath("RYYIYR4").Value != "NO MESSAGE LEFT" delim noteGui["e7"].Value
LatitudeEl.ElementFromPath("RYYIYR4").Value := "NO MESSAGE LEFT" delim noteGui["e7"].Value
SendInput "^{END}"
   }
  }
 } ;try
catch as e{
MsgBox "There was an error processing your request. If the issue persists, please restart NoteAssist.","Error 0x" e.line,16+262144
} ;catch
} ;win ex
Else
MsgBox "Latitude not found","Error",48+262144
}

/*************
*    WN     *
*************/
sbclick017(Button, Info){
if WinExist("Latitude - Account Work Form"){
try{
obcall()
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
if LatitudeEl.ElementFromPathExist("RYYIYR2q"){
LatitudeEl.ElementFromPathExist("RYYIYR2q").ToggleState := "1"
}
if LatitudeEl.ElementFromPathExist("RYYIYR3/K"){
if LatitudeEl.ElementFromPath("RYYIYR3/K").Type = 50020{
if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "WN - WRONG/NO GOOD NUMBER"
LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "WN - WRONG/NO GOOD NUMBER"
if LatitudeEl.ElementFromPath("RYYIYR4").Value != "WRONG NUMBER" delim noteGui["e7"].Value
LatitudeEl.ElementFromPath("RYYIYR4").Value := "WRONG NUMBER" delim noteGui["e7"].Value
SendInput "^{END}"
   }
  }
 } ;try
catch as e{
MsgBox "There was an error processing your request. If the issue persists, please restart NoteAssist.","Error 0x" e.line,16+262144
}
} ;winex
Else
MsgBox "Latitude not found","Error",48+262144
}

/*************
*    TRC     *
*************/
sbclick018(Button, Info){
if WinExist("Latitude - Account Work Form"){
try{
inbcall()
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
if LatitudeEl.ElementFromPathExist("RYYIYR3/K"){
if LatitudeEl.ElementFromPath("RYYIYR3/K").Type = 50020{
if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "TRC" delim "TRANSFERRED CALL"
LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "TRC" delim "TRANSFERRED CALL"
if LatitudeEl.ElementFromPath("RYYIYR4").Value != "TRANSFERRED CALL TO: " noteGui["e7"].Value
LatitudeEl.ElementFromPath("RYYIYR4").Value := "TRANSFERRED CALL TO: " noteGui["e7"].Value
SendInput "^{END}"
   }
  }
 } ;try
catch as e{
MsgBox "There was an error processing your request. If the issue persists, please restart NoteAssist.","Error 0x" e.line,16+262144
}

} ;winex
Else
MsgBox "Latitude not found","Error",48+262144
}

/*************
*     AML    *
*************/
sbclick019(Button, Info){
if WinExist("Latitude - Account Work Form"){
try{
mancall()
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
if LatitudeEl.ElementFromPathExist("RYYIYR3/K"){
if LatitudeEl.ElementFromPath("RYYIYR3/K").Type = 50020{
if LatitudeEl.ElementFromPath("RYYIYR3/K").Value != "AML - LEFT MESSAGE ANSWERING MACHINE"
LatitudeEl.ElementFromPath("RYYIYR3/K").Value := "AML - LEFT MESSAGE ANSWERING MACHINE"
if LatitudeEl.ElementFromPath("RYYIYR4").Value != "AML" delim noteGui["e7"].Value
LatitudeEl.ElementFromPath("RYYIYR4").Value := "AML" delim noteGui["e7"].Value
SendInput "^{END}"
   }
  }
 } ;try
catch as e{
MsgBox "There was an error processing your request. If the issue persists, please restart NoteAssist.","Error 0x" e.line,16+262144
}
} ;winex
Else
MsgBox "Latitude not found","Error",48+262144
}
/*********************
*    WORKING PQ      *
*********************/
sbclick020(Button, Info){
PIFLRres := MsgBox("Would you like to prepare a Working PQ note?","Post PIFLR?",32+262144 " YesNo")
If PIFLRres = "No"
	Return
Else
	If WinExist("Latitude - Account Work Form"){
		WinActivate "Latitude - Account Work Form"
		SendInput "!n"
		sleep 250
		SendInput "!n"
		WinWait "Latitude - New Note"
		If WinExist("Latitude - New Note"){
		WinActivate "Latitude - New Note"
		SendInput "CO{Down}{Tab}CO{Down}{Tab}"
		SendInput "WORKING PQ - "
		}
	}
	
	else	
		MsgBox "Latitude not found","Error",48+262144
}

/****************
*    atynote    *
****************/

sbclick021(Button, Info){ ;aty
notegui["atytitle"].Visible := 1
loop 4{
notegui["atyt" A_index].Visible := 1
notegui["atye" A_index].Value   := ""
notegui["atye" A_index].Visible := 1
}
notegui["atyb1"].Visible := 1
notegui["atyb2"].Visible := 1

		If noteGui["slepaTitle"].Visible = 1{
		Result := MsgBox("Would you like to close the SLEPA?","Close SLEPA?",32+262144 " YesNo")
		if Result = "No"
		Return
		Else
		CloseSLEPA()
		}		
		
		If noteGui["EPATitle"].Visible = 1{
		Result := MsgBox("Would you like to close the EPA?","Close?",32+262144 " YesNo")
		if Result = "No"
			Return
		else
		CloseEPA()
		}

		If noteGui["FCTitle"].Visible = 1{
		Result := MsgBox("Would you like to close the F/C?","Close?",32+262144 " YesNo")
		if Result = "No"
			Return
		else		
		CloseFC()
		}	

For bn, label in noteBSCMap{
noteGui["sb" bn].Visible := 0		
}
noteGui["gb1"].Visible := 0
noteGui["gb2"].Visible := 0

}

/**************
*  BUTTON 999 *
**************/
sbclick999(Button, Info){
}




/*************
*  MENU BAR  *
*************/

	noteMenu := MenuBar()															;Make a Menu Bar
	
	ToolMenu := Menu()
	MassUpdaters := Menu()
	ToolMenu.Add "&PTPIF", (*) => ptpif()
	ToolMenu.Add "&SIF Calculator", (*) => CalcOpen()
	MassUpdaters.Add "&PIFLR All", (*) => piflrall()
	MassUpdaters.Add "&CO/CANTC All", (*) => cantcall()
	MassUpdaters.Add "&Cease Comm All", (*) => ceaseall()
	MassUpdaters.Add "&Remove Cease Comm All"  , (*) => ceaseremall()
	ToolMenu.Add("Mass Updaters", MassUpdaters)
	noteMenu.Add "&Tools", ToolMenu
	
	HelpMenu := Menu()																;Make a Help Menu Item
	HelpMenu.Add "&Settings", (*) => SettingsGUI()									;Add a Settings option to HelpMenu Item
	HelpMenu.Add "&About", (*) => MsgBox("Version " Version "`njcampbell@pfcusa.com","About",262144) ;Add an About option to HelpMenu Item
	HelpMenu.Add "&Reload", (*) => Reload()											;Add a Reload option to HelpMenu Item
	noteMenu.Add "&Help", HelpMenu													;Add Menu Item to Menu Bar

	
	
	noteGui.MenuBar := noteMenu														;Set GUI Menu Bar to our Menu Bar
	
	
/* End */ ;close end finish wrap done finito
	epahas := 0
	for val, key in epacMap{
		noteGui["epae" val].Visible := epahas
		noteGui["epat" val].Visible := epahas
		noteGui["epae" val].Value 	:= ""
	}
	noteGui["tN"].Visible 		:= epahas
	noteGui["tY"].Visible 		:= epahas
	noteGui["cbOK1"].Visible 	:= epahas
	noteGui["cbOK2"].Visible 	:= epahas
	noteGui["epadd"].Visible 	:= epahas
	noteGui["cbRec"].Visible 	:= epahas
	noteGui["EPATitle"].Visible := epahas
	noteGui["epab1"].Visible 	:= epahas
	noteGui["epab2"].Visible 	:= epahas
	noteGui["epab3"].Visible 	:= epahas
	noteGui["epab4"].Visible 	:= epahas
	noteGui["epab5"].Visible 	:= epahas
	noteGui["epab6"].Visible 	:= epahas
	
	noteGui["cbOK1"].Text 		:= ""
	noteGui["cbOK2"].Text 		:= ""
	noteGui["cbRec"].Text 		:= ""
	noteGui["epadd"].Text		:= ""
	noteGui["epadd"].Enabled 	:= epahas
	
	For key, val in fccMap{
	; noteGui["fce" key].Visible := 0
	; noteGui["fce" key].Value := ""
	noteGui["fct" key].Visible := 0
	}
	For key, val in fccbMap{
	noteGui[key].Visible := 0
	noteGui["t" key].Visible := 0
	}
	noteGui["FCTitle"].Visible := 0
	noteGui["fcb1"].Visible := 0
	noteGui["fcb2"].Visible := 0
	noteGui["fcb3"].Visible := 0
	noteGui["fceMS"].Visible := 0
	; noteGui["fceRO"].Visible := 0
	noteGui["fceEM"].Visible := 0
	noteGui["fceBN"].Visible := 0
	
	For key, val in slepaMap{
	noteGui["slepae" key].Visible := 0
	noteGui["slepae" key].Value := ""
	noteGui["slepat" key].Visible := 0
	}
	noteGui["slepaTitle"].Visible := 0
	noteGui["slepab1"].Visible := 0
	noteGui["slepab2"].Visible := 0
	noteGui["slepab3"].Visible := 0
	noteGui["slepab4"].Visible := 0
	
	noteGui["CalcTitle"].Visible 		:= 0
	loop 4{
	 notegui["calct" A_index].Visible 	:= 0
	 notegui["calce" A_index].Visible 	:= 0
	}
	notegui["calce5"].Visible 			:= 0
	notegui["discUD"].Visible 			:= 0
	notegui["calcb1"].Visible 			:= 0	
	
	noteGui["gb1"].Visible := 0
	noteGui["gb2"].Visible := 0
	
	notegui["atytitle"].Visible := 0
	loop 4{
	notegui["atyt" A_index].Visible := 0
	notegui["atye" A_index].Value   := ""
	notegui["atye" A_index].Visible := 0
	}
	notegui["atyb1"].Visible := 0
	notegui["atyb2"].Visible := 0
	
	noteGui.Hide
	noteGui.Show("Autosize")

} ;END NOTEGUI

noteGui.OnEvent("Close", CloseNotes)

CloseNotes(GuiObject){

	ResultClose := MsgBox("Would you like to close Note Assist?","Close?",32+262144 " YesNo")
	if ResultClose = "No"
		return
	else{
		WinKill
		;WinKill "NoteAssist.exe"
		;WinKill "NoteAssistTesting.exe"
		exitapp()
	}
}
	
	noteGui.Show("AutoSize")


SettingsGUI(){

;DetectHiddenWindows True
; WinGetPos &nax, &nay,,, "Note Assist Testing Environment"
if WinExist("Note Assist Settings"){
	WinShow "Note Assist Settings"
	WinActivate "Note Assist Settings"
	; WinMove nax, nay,,,"Note Assist Settings"
	}
else{
Global setGui := gui("AlwaysOnTop","Note Assist Settings")

	;Variables

		setGui.MarginX := setGui.MarginY := Margin/1.5
		settw   := Margin*8.5
		setth   := Margin*2
		setew   := Margin*20
		seteh   := Margin*2
		sbtnw   := Margin*9
		sbtnh   := Margin*4
		setw    := Margin*2+(settw+setth)
		seth    := Margin*2
		noteGui.GetPos(&rx,&ry,&rw,&rh)
		noteGuiw := rw+Margin*2

		;Maps
		
			Global setcmap := Map(
			"1",	"Font Size",
			; "2",	"",
			; "3",	"",
			; "4",	"",
			; "5",	"",
			; "6",	"",
			; "7",	"",
			; "8",	"",
			; "9",	"",
			)		
	
	;GUI Options
		setGui.SetFont(settings["str"] " s" Margin " c" settings["fontcolor"], settings["name"])
		setGui.Title := "Note Assist Settings"
		setGui.BackColor := settings["cc"]

	;Controls
	
		;MainWindow
		
			/*Font Size and Name*/
			
			setGui.Add("Text","xm w" settw+15 " h" setth " 0x200 vt1", "Font Size")
			setGui.Add("Edit","xp+" settw+15 " w" setew+15 " h" seteh " ReadOnly Uppercase 0x200 ve1","" settings["size"])
			setGui.Add("UpDown", "vsetUD Range6-32", settings["size"])
			setGui.Add("Text","xm w" settw+15 " h" setth " 0x200 vt2", "Font Name")
			setGui.Add("Edit","xp+" settw+15 " w" setew+15 " h" seteh " ReadOnly Uppercase background" settings["cc"] " 0x200 ve2","" settings["name"])
			setGui["e2"].SetFont(,settings["name"])
			
			/*Delimiter*/
			
			setGui.Add("Text","xm w" settw+15 " h" setth " 0x200 vt3", "Delimiter")
			setGui.Add("Text","xm+" settw+15 " yp+" Margin/4 " w" seteh*0.75 " h" seteh*0.75 " 0x200 vcbdash Center BackgroundWhite border","")
			setGui.Add("Text","xp+" nteh*0.90+15 " yp" " h" nteh*0.75 " vtdash","Dash (-)")
			setGui.Add("Text","xp+" settw " w" seteh*0.75 " h" seteh*0.75 " 0x200 vcbcomma Center BackgroundWhite border","")
			setGui.Add("Text","xp+" seteh*0.90 " h" seteh*0.75 " vtcomma","Comma (,)")
			
			/*Date Format*/
			
			setGui.Add("Text","xm w" settw+15 " h" setth " 0x200 vt4", "Date Format")
			setGui.Add("Text","xm+" settw+15 " yp+" Margin/4 " w" seteh*0.75 " h" seteh*0.75 " 0x200 vcbdf Center BackgroundWhite border","")
			setGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtyf","YY-MM-DD")
			
			setGui.Add("Text","xp+" settw+15 " w" seteh*0.75 " h" seteh*0.75 " 0x200 vcbmf Center BackgroundWhite border","")
			setGui.Add("Text","xp+" seteh*0.90 " h" seteh*0.75 " vtmf","MM-DD-YY")
			
			/*Department*/
			
			setGui.Add("Text","xm w" settw+15 " h" setth " 0x200 vt5", "Department")
			setGui.Add("Text","xm+" settw+15 " yp+" Margin/4 " w" seteh*0.75 " h" seteh*0.75 " 0x200 vcbls1 Center BackgroundWhite border","")
			setGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtls1","Early Stage")
			
			setGui.Add("Text","xp+" settw+15 " w" seteh*0.75 " h" seteh*0.75 " 0x200 vcbls2 Center BackgroundWhite border","")
			setGui.Add("Text","xp+" seteh*0.90 " h" seteh*0.75 " vtls2","Late Stage")
			
			/*Always On Top*/
			setGui.Add("Text","xm w" settw+15 " h" setth " 0x200 vt6", "Always On Top")
			setGui.Add("Text","xm+" settw+15 " yp+" Margin/4 " w" seteh*0.75 " h" seteh*0.75 " 0x200 vcbaot1 Center BackgroundWhite border","")
			setGui.Add("Text","xp+" nteh*0.90 " h" nteh*0.75 " vtaot1","Yes")
			
			setGui.Add("Text","xp+" settw+15 " w" seteh*0.75 " h" seteh*0.75 " 0x200 vcbaot2 Center BackgroundWhite border","")
			setGui.Add("Text","xp+" seteh*0.90 " h" seteh*0.75 " vtaot2","No")
			
	;Buttons

			setGui.Add("Button","xm yp+" 150 " w" sbtnw " h" sbtnh-10,"Close").OnEvent("Click", SetClose)
			setGui.Add("Button","xp+" sbtnw+Margin+9 " w" sbtnw " h" sbtnh-10,"Change Font").OnEvent("Click", fontsender)
			setGui.Add("Button","xp+" sbtnw+Margin+9 " yp w" sbtnw " h" sbtnh-10,"Apply").OnEvent("Click", SetSave)
			

		;Button Functions
		
		fontsender(Button, Info){ ;This was really such a pain to make but it is working so...
			If (!isSet(settings))
				settings := ""
	
			settings := FontSelect(settings,noteGui.hwnd) ; shows the user the font selection dialog
	
			If (!settings)
				return ; to get info from settings use ... bold := settings["bold"], settings["name"], etc.
			
			setGui["e2"].SetFont(settings["str"],settings["name"])
			setGui["e2"].text := settings["name"]
			fstr := settings["str"]
			fsize := ""
			RegExMatch(fstr, "(s+\d+)", &fsize)
			RegWrite fsize[],"REG_SZ","HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "size"
			;RegWrite fstr,"REG_SZ","HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "str"
			setGui["e1"].Value := settings["size"]
			
			; ctl.gui["MyEdit1"].SetFont(settings["str"],settings["name"]) ; apply font+style in one line, or...
			; ctl.gui["MyEdit2"].SetFont(settings["str"],settings["name"])
		}
		
		SetSave(*){ ;'Apply' Button Function
		
		Result := MsgBox("Applying will clear your notes. Would you like to continue?","Apply Settings?",16+262144 " YesNo")
		if Result = "No"
			Return
		else
		
		;RegWrite val, "REG_SZ", "HKEY_CURRENT_USER\SOFTWARE\NoteAssist", header
		;RegWrite Value, ValueType, KeyName [, ValueName]
		
			RegWrite setGui["e1"].Value,"REG_SZ","HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "size"
			RegWrite setGui["e2"].Value,"REG_SZ","HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "name"
			
			If setGui["cbdash"].text = checkmark
				RegWrite ' - ',"REG_SZ","HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "delim"
			
			If setGui["cbcomma"].text = checkmark
				RegWrite ', ',"REG_SZ","HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "delim"
			
			If setGui["cbdf"].Text = checkmark
				RegWrite 'yyyy-MM-dd' ,"REG_SZ","HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "datefrmt"
			
			If setGui["cbmf"].Text = checkmark
				RegWrite 'MM-dd-yyyy' ,"REG_SZ","HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "datefrmt"
			
			If setGui["cbls1"].Text = checkmark
				RegWrite 0 ,"REG_SZ","HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "esls"
			
			If setGui["cbls2"].Text = checkmark
				RegWrite 1 ,"REG_SZ","HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "esls"				
			
			If setGui["cbaot1"].Text = checkmark
				RegWrite 1 ,"REG_SZ","HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "aaot"
			
			If setGui["cbaot2"].Text = checkmark
				RegWrite 0 ,"REG_SZ","HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "aaot"
			
			Reload
		}

	;Other Controls

	;Other Control Functions
	
		;Checkbox Func
		if settings["delim"] = " - "
			setGui["cbdash"].Text := checkmark
		
		if settings["delim"] = ", "
			setGui["cbcomma"].Text := checkmark
			
		if settings["datefrmt"] = "yyyy-MM-dd"
			setGui["cbdf"].Text := checkmark
		
		if settings["datefrmt"] = "MM-dd-yyyy"
			setGui["cbmf"].Text := checkmark
			
		if settings["esls"] = 0
			setGui["cbls1"].Text := checkmark

		if settings["esls"] = 1
			setGui["cbls2"].Text := checkmark	

		if settings["aaot"] = 0
			setGui["cbaot2"].Text := checkmark

		if settings["aaot"] = 1
			setGui["cbaot1"].Text := checkmark			
			
				
		setGui["tdash"].OnEvent("Click", cbDashc)
		setGui["cbdash"].OnEvent("Click", cbDashc)
			cbDashc(Text, Info){
			If setGui["cbdash"].Text = checkmark
				setGui["cbdash"].Text := ""
			Else If setGui["cbdash"].Text = ""
				setGui["cbdash"].Text := checkmark
				setGui["cbcomma"].Text := ""
			} ;End cbDashc
			
		setGui["tcomma"].OnEvent("Click", cbCommac)
		setGui["cbcomma"].OnEvent("Click", cbCommac)
			cbCommac(Text, Info){
			If setGui["cbcomma"].Text = checkmark
				setGui["cbcomma"].Text := ""
			Else If setGui["cbcomma"].Text = ""
				setGui["cbcomma"].Text := checkmark
				setGui["cbdash"].Text := ""
			} ;End cbCommac
			
		setGui["tyf"].OnEvent("Click", cbyfc)
		setGui["cbdf"].OnEvent("Click", cbyfc)
			cbyfc(Text, Info){
			If setGui["cbdf"].Text = checkmark
				setGui["cbdf"].Text := ""
			Else If setGui["cbdf"].Text = ""
				setGui["cbdf"].Text := checkmark
				setGui["cbmf"].Text := ""
			} ;End cbDashc	

		setGui["tmf"].OnEvent("Click", cbmfc)
		setGui["cbmf"].OnEvent("Click", cbmfc)
			cbmfc(Text, Info){
			If setGui["cbmf"].Text = checkmark
				setGui["cbmf"].Text := ""
			Else If setGui["cbmf"].Text = ""
				setGui["cbmf"].Text := checkmark
				setGui["cbdf"].Text := ""
			} ;End cbDashc


		setGui["cbls1"].OnEvent("Click", cblate0)
		setGui["tls1"].OnEvent("Click", cblate0)
			cblate0(Text, Info){
			If setGui["cbls1"].Text = checkmark
				setGui["cbls1"].Text := ""
			Else If setGui["cbls1"].Text = ""
				setGui["cbls1"].Text := checkmark
				setGui["cbls2"].Text := ""
			} ;End cblate0
			
		setGui["cbls2"].OnEvent("Click", cblate1)
		setGui["tls2"].OnEvent("Click", cblate1)
			cblate1(Text, Info){
			If setGui["cbls2"].Text = checkmark
				setGui["cbls2"].Text := ""
			Else If setGui["cbls2"].Text = ""
				setGui["cbls2"].Text := checkmark
				setGui["cbls1"].Text := ""
			} ;End cblate1


		setGui["cbaot1"].OnEvent("Click", cbaot0)
		setGui["taot1"].OnEvent("Click", cbaot0)
			cbaot0(Text, Info){
			If setGui["cbaot1"].Text = checkmark
				setGui["cbaot1"].Text := ""
			Else If setGui["cbaot1"].Text = ""
				setGui["cbaot1"].Text := checkmark
				setGui["cbaot2"].Text := ""
			} ;End cbaot0
			
		setGui["cbaot2"].OnEvent("Click", cbaot1)
		setGui["taot2"].OnEvent("Click", cbaot1)
			cbaot1(Text, Info){
			If setGui["cbaot2"].Text = checkmark
				setGui["cbaot2"].Text := ""
			Else If setGui["cbaot2"].Text = ""
				setGui["cbaot2"].Text := checkmark
				setGui["cbaot1"].Text := ""
			} ;End cbaot1
	
	;Misc Functions
		
		SetClose(*){
			setGui.Hide
		}



	;GUI Display	
	
	setGui.Show("AutoSize")
  } ;End Else
  
}

; noteRegWrite()

; noteRegWrite(){

; /************************************************************
; * This is function establishes a current_user_registry		*
; * It only does so if there has been no local registry 		*
; * that has been made before. This way settings can be 		*
; * saved (locally) to the registry without necessitating		*
; * the need for additional files that will go along with the	*
; * program.													*
; *************************************************************/

; TRY RegRead("HKEY_CURRENT_USER\SOFTWARE\NoteAssist", "rbf") = 1 	;Reads the current_user_registry
; IF A_LastError = 2{ 												;If windows returns a 'no file found error' then writes the following to the registry
; global settings := Map(
	; "name"		,""	,
	; "size"		,"10"				,
	; "color"		,""					,
	; "fontcolor"	,000000				,
	; "strike"	,1					,
	; "underline"	,1					,
	; "italic"	,1					,
	; "bold"		,1					,
	; "cc"		,""					,
	; "bgcolor"	,"F5F5F5"			,
	; "str"		,"norm s10.0"		,
	; "delim"		," - "				,
	; "datefrmt"	,"MM-dd-yyyy"		,
	; "rbf"		,"1"				,
	; "esls"		,0					,
	; "aaot"		,0					,
; )
; for header, val in settings
; RegWrite val, "REG_SZ", "HKEY_CURRENT_USER\SOFTWARE\NoteAssist", header
; }
; ELSE
; RETURN
; }

/**************************************
* this is a posting note Calculator   *
* made for First agents posting notes *
**************************************/

; postcalc()

postcalc(){

postcalc := gui(,"Posting Note Calculator")

Margin := 10
margsm := Margin-Margin/1.2
postcalc.MarginX := postcalc.MarginY := Margin-Margin/1.2
	Global nttw    	:= Margin*8
	Global ntth    	:= Margin*2
	Global ntew    	:= Margin*20
	Global nteh    	:= Margin*2
	Global btnw   	:= Margin*9
	Global btnh   	:= Margin*4
	Global ntw     	:= Margin*2
	Global nth     	:= Margin*2
	
postcalc.SetFont("s" Margin)	

Global postcalctm := Map(
"01", "File 1:",
"02", "File 2:",
"03", "File 3:",
"04", "File 4:",
"05", "File 5:",
"06", "File 6:",
"07", "File 7:",
"08", "File 8:",
"09", "File 9:",
"10", "File 10:",
"11", "File 11:",
"12", "File 12:",
"13", "File 13:",
"14", "File 14:",
"15", "File 15:"
)

postcalc.Add("Text","xm w" nttw*2 " h" ntth " border center 0x200 vpctfilt", "File Number")
postcalc.Add("Text","yp w" nttw*2-70 " h" ntth " border center 0x200 vpctbalt", "Balance")

for key, val in postcalctm{
postcalc.Add("Text","xm w" nttw-35 " h" ntth " 0x200 vpct" key, val)
postcalc.Add("Edit","xp+" nttw-35 " w" nttw+35 " yp h" ntth " right vpce" A_Index).OnEvent("Change", bbalupd)

postcalc.Add("Text","yp w" nttw-70 " h" ntth " 0x200 vpct1" key, "$")
postcalc.Add("Edit","xp+" nttw-70 " w" nttw " yp h" ntth " right vpcev" A_Index,"0.00").OnEvent("Change", bbalupd)
}

;nttw-35+nttw+70+nttw-70+nttw+70
;nttw*4+35

postcalc["pctbalt"].GetPos(&x1, &y1, &w1, &h1)
postcalc.Add("Text","xm w" nttw+35 " h" ntth " 0x200 vpctbbal", "Beginning Balance:")
postcalc.Add("Edit","x" x1+nttw-70 " w" nttw " yp h" ntth " ReadOnly right vpcebbal","0.00").OnEvent("Change", bbalupd)
postcalc.Add("Text","xm w" nttw+35 " h" ntth " 0x200 vpctdisc", "Discount Amount %",)
postcalc.Add("Edit","x" x1+nttw-70 " w" nttw-35 " yp h" ntth " right vpcedisc","0").OnEvent("Change", bbalupd)
postcalc.Add("UpDown", "vsetUD Range0-100","0")
postcalc["pcedisc"].OnEvent("LoseFocus",adhd)
adhd(Edit, Info){
if !IsNumber(postcalc["pcedisc"].value)
postcalc["pcedisc"].value := 0
}

postcalc["pce1"].GetPos(&x2, &y2, &w2, &h2)
postcalc.Add("Text","x" x1+nttw*2-70+margsm " y" y1 " w" nttw*4-76 " h" ntth " border center 0x200 vpostcalct", "Payment Amount To Each File")

for key, val in postcalctm{
postcalc["pce" A_Index].GetPos(&x2, &y2, &w2, &h2)
postcalc.Add("Text","x" x1+nttw*2-70+margsm " y" y2 " w" nttw-35 " h" ntth " 0x200 vpctt" key, "|" val)
postcalc.Add("Edit","xp+" nttw-33 " w" nttw+25 " yp h" ntth " right ReadOnly vpcee" A_Index)

postcalc.Add("Text","yp w" nttw-70 " h" ntth " 0x200 vpctt1" key, "$")
postcalc.Add("Edit","xp+" nttw-70 " w" nttw " yp h" ntth " right ReadOnly vpceev" A_Index)
}

postcalc["pcebbal"].GetPos(&x3, &y3, &w3, &h3)
postcalc.Add("Text","x" x3+w3+margsm " y" y3 " w" nttw+35 " h" ntth " center 0x200 vpctbbal1", "Total Payments:")
postcalc["pceev15"].GetPos(&x3, &y3, &w3, &h3)
postcalc["pctbbal1"].GetPos(&x4, &y4, &w4, &h4)
postcalc.Add("Edit","x" x3 " y" y4 " w" nttw " h" ntth " ReadOnly right vpcebbal1")

postcalc["pcebbal"].GetPos(&x5, &y5, &w5, &h5)
postcalc.Add("Text","xm w" nttw+35 " h" ntth " 0x200 vpostnotet", "Posting Note:")
postcalc.Add("Edit","x" x1+nttw-70 " w" nttw*5-76 " yp h" ntth*3 " readonly vpostnotee")

postcalc.Add("Button","xm w" btnw " h" btnh " vclosepcb", "Close").OnEvent("Click", ClosePNC)

ClosePNC(Button, Info){
Result := MsgBox("Would you like to close the calculator?","Close?",262144 " YesNo")
if Result = "No"
Return
else{
WinKill "Posting Note Calculator"
ExitApp
 }
}

postcalc.Add("Button","xp+" x5+nttw*5-77-btnw " w" btnw " h" btnh " vcopypcb", "Copy").OnEvent("Click", CopyPNC)

CopyPNC(Button, Info){
A_Clipboard := postcalc["postnotee"].value
}

postcalc.Add("Button","x" x5+nttw*5-77-(btnw*2)-margsm " yp w" btnw " h" btnh " vclearpcb", "Clear").OnEvent("Click", ClearPNC)

ClearPNC(Button, Info){
Result := MsgBox("Clear the calculator?","Calculator Clear?",262144 " YesNo")
if Result = "No"
Return
else{
postcalc["pcedisc"].value := 0
pcdisc := postcalc["pcedisc"].value/100
filestr := ""
for key, val in postcalctm{
postcalc["pce" A_Index].value := ""
postcalc["pcee" A_Index].value := postcalc["pce" A_Index].value
postcalc["pcev" A_Index].value := "0.00"
; postcalc["pceev" A_Index].value := Round(postcalc["pcev" A_Index].value-(postcalc["pcev" A_Index].value*pcdisc),2)
if postcalc["pceev" A_Index].value < 0{
postcalc["pceev" A_Index].value := "(" postcalc["pceev" A_Index].value ")"
}
if postcalc["pce" A_Index].value = ""
filestr := filestr
else
filestr .= "File " postcalc["pce" A_Index].value ": " postcalc["pceev" A_Index].value " - "
bbalupd(1,1)
idnm := 0
leftpnc := 0
  }
 }
}

ClearPNC1(){
postcalc["pcedisc"].value := 0
pcdisc := postcalc["pcedisc"].value/100
filestr := ""
for key, val in postcalctm{
postcalc["pce" A_Index].value := ""
postcalc["pcee" A_Index].value := postcalc["pce" A_Index].value
postcalc["pcev" A_Index].value := "0.00"
bbalupd(1,1)
  }
 }

postcalc.Add("Button","x" x5+nttw*5-80-(btnw*3)-margsm " yp w" btnw " h" btnh " vgathercb", "Import").OnEvent("Click", SumPNC)

global idnm := 0
global leftpnc := 0

SumPNC(Button, Info){
Result := MsgBox("Would you like to import the file balances?","Import",262144 " YesNo")
if Result = "No"
Return
else{
SetTitleMatchMode 2
if WinExist("Latitude - Account Work Form"){
try{
LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
WinGetPos &latx, &laty,,, "Latitude - Account Work Form"
if LatitudeEl.ElementFromPathExist("Yw0"){
LatitudeEl.ElementFromPath("Yw0").Invoke()
SetTitleMatchMode "RegEx"
WinWait "Link \d{4,8}"
WinActivate "Link \d{4,8}"
WinMove latx, laty,,, "Link \d{4,8}"
LinkEl := UIA.ElementFromHandle("Link \d{4,8}")
numfiles := LinkEl.ElementFromPath("YYqNO").Children
if leftpnc = 0
leftpnc := numfiles.length - idnm
if leftpnc > 15
leftpnc := 15
Result := MsgBox("Get the next " leftpnc " files?","Calculator Clear?",262144 " YesNo")
if Result = "No"
Return
else{
ClearPNC1()
rowfill := 1
if numfiles.length - idnm <= 0{
leftpnc := 0
idnm := 0
}
loop numfiles.length - idnm{
;MsgBox idnm " & " A_index
If A_index > 15{
global leftpnc := numfiles.length - idnm
if leftpnc > 15
leftpnc := 15
MsgBox "Only able to load 15 accounts automatically. Press 'Import' again to get the next " leftpnc " files. Pressing 'Clear' will reset this counter."
SetTitleMatchMode 2
break
}
idnm++
LinkEl.WalkTree("1,2,1,1," idnm).Select()
LinkEl.ElementFromPath("1,1,1,1").Select()
datev := UIA.ElementFromHandle("Link \d{4,8}")
if datev.ElementFromPath("1,6,1,1,6").Value = ""{
acctnum := datev.ElementFromPath("1,2,1,1," idnm).Value
acctbal := datev.ElementFromPath("1,6,1,3,33").Value
postcalc["pce" rowfill].value	:= acctnum
postcalc["pcev" rowfill].value 	:= RegExReplace(acctbal,"[^0-9.]","")
rowfill++
bbalupd(1,1)
}
Else{
rowfill := rowfill
}
bbalupd(1,1)
If numfiles.length - idnm = 0{
MsgBox "Done"
leftpnc := 0
idnm := 0
bbalupd(1,1)
     }
    }
   }
  }
  Else{
  MsgBox "No links found."
  SetTitleMatchMode 2
  WinActivate "Latitude - Account Work Form"
  SendInput "!illlll{Enter}"
  sleep 250
  MenuEl := UIA.ElementFromHandle("Latitude - Account Work Form")
  postcalc["pce1"].value := MenuEl.ElementFromPath("RYYbQU").Value
  MenuEl.ElementFromPathExist("RYYbR/E").Value := 435
  postcalc["pcev1"].value := MenuEl.ElementFromPath("RYYbQUz").Value
  MenuEl.ElementFromPathExist("RYYbR/E").Value := 0
  }
 }
}
Else
MsgBox "Latitude not found","Error",262144 
}
}


;"Beginning Balance" $2,681.83 Discount Amount% 50 ($1,340.92), Payment Amount to Each File: 14967760-$179.92, File: 0-$1,161.00, File: 0-$0.00, File: 0-$0.00, File: 0-$0.00, File: 0-$0.00, File: 0-$0.00, File: 0-$0.00, File: 0-$0.00, File: 0-$0.00, File: 0-$0.00, File: 0-$0.00, File: 0-$0.00, File: 0-$0.00, File: 0-$0.00 Total Payment $1,340.92													
;  ($142.18), File: 157267849 - $369.60, File: 157267850 - $673.03, Total Payment $1042.63													

bbalupd(Edit, Info){


pcdisc := postcalc["pcedisc"].value
if !isnumber(postcalc["pcedisc"].value){
pcdisc := 0
pcdisc := pcdisc/100
}
Else 
pcdisc := postcalc["pcedisc"].value/100

filestr := ""
valstr := ""
balpc := 0
balpc1 := 0

for key, val in postcalctm{
col2val := postcalc["pcev" A_Index].value
if !IsNumber(col2val)
valstr := "0.00"
else
valstr := postcalc["pcev" A_Index].value ;valstr = the index of the first ammt boxes

postcalc["pceev" A_Index].value := valstr ;the second row index location is given our valstr value

;postcalc["pceev" A_Index].value := Round(valstr-(valstr*pcdisc),2) ;eg the rounded value of 100 - (100*10%)

row2val := valstr*pcdisc ; 13.33*0.1 = 1.333
row2val := floor(row2val*100) ;round up to the next int on 1.333*100 = 133.3 rounded to the next int is 133
row2val := row2val/100 ;133/100 = 1.33
row2val := valstr-row2val
;MsgBox row2val
postcalc["pceev" A_Index].value := Round(row2val,2) ;now our row values are all rounded up from our discount

		; discbal := pri*(disc/100)
		; discbal := floor(discbal*100)
		; discbal := discbal/100
		; finalbal := pri-discbal

if postcalc["pce" A_Index].value = ""
filestr := filestr
else
filestr .= "File " postcalc["pce" A_Index].value ": $" postcalc["pceev" A_Index].value " - "
postcalc["pcee" A_Index].value := postcalc["pce" A_Index].value

try{
balpc 	+= postcalc["pcev" A_Index].value ;account totals
balpc1 	+= postcalc["pceev" A_Index].value ;rounded up row 2 totals
}

		; discbal := pri*(disc/100)
		; discbal := floor(discbal*100)
		; discbal := discbal/100
		; finalbal := pri-discbal



postcalc["pcebbal"].value 	:= Round(balpc,2)
postcalc["pcebbal1"].value 	:= Round(balpc1,2)

if postcalc["pceev" A_Index].value < 0{
postcalc["pceev" A_Index].value := "(" postcalc["pceev" A_Index].value ")"
 }
}

if balpc1 < 0{
postcalc["pcebbal1"].value := "(" balpc1 ")"
}

totdisc := Round(balpc, 2) - Round(balpc1, 2)
totdisc := Round(totdisc,2)
postcalc["postnotee"].value := "Beginning Balance: $" postcalc["pcebbal"].Value " - Discount Amount: " postcalc["pcedisc"].Value "% ($" totdisc ") - Payment Amount to Each File: " filestr "Total Payment: $" postcalc["pcebbal1"].Value
 }
bbalupd(1,1)
postcalc.Show("Autosize")
}



/******************************************************************************
*   Another attempt at DPI adjustment - this function is passed information   *
*   from the window parameter and uses a WindowsDLL to try and                *
*   parse this information to an address and return a value                   *
******************************************************************************/

;Click(67*WinGetDpi("A")/144, 494*WinGetDpi("A")/144)

WinGetDpi(WinTitle?, WinText?, ExcludeTitle?, ExcludeText?) {
    return (hMonitor := DllCall("MonitorFromWindow", "ptr", WinExist(WinTitle?, WinText?, ExcludeTitle?, ExcludeText?), "int", 2, "ptr") ; MONITOR_DEFAULTTONEAREST
    , DllCall("Shcore.dll\GetDpiForMonitor", "ptr", hMonitor, "int", 0, "uint*", &dpiX:=0, "uint*", &dpiY:=0), dpiX)
}

/************************************
*   Basic range function   			*
*   useful but currently unused   	*
************************************/



range(a, b:=unset, c:=unset) {
   IsSet(b) ? '' : (b := a, a := 0)
   IsSet(c) ? '' : (a < b ? c := 1 : c := -1)

   pos := a < b && c > 0
   neg := a > b && c < 0
   if !(pos || neg)
      throw Error("Invalid range.")

   return (&n) => (
      n := a, a += c,
      (pos && n < b) OR (neg && n > b) 
   )
}

/*********************************************
* This is not used, but would allow us to    *
* to make the background a different color.  *
*********************************************/  

; /*
; ColorSelect(Color := 0, hwnd := 0, &custColorObj := "",disp:=false) { ;A function to call the windows built in color DLL
    ; Static p := A_PtrSize
    ; disp := disp ? 0x3 : 0x1 ; init disp (0x3 = full panel / 0x1 = basic panel)
    
    ; If (custColorObj.Length > 16)
        ; throw Error("Too many custom colors.  The maximum allowed values is 16.")
    
    ; Loop (16 - custColorObj.Length)
        ; custColorObj.Push(0) ; fill out custColorObj to 16 values
    
    ; CUSTOM := Buffer(16 * 4, 0) ; init custom colors obj
    ; CHOOSECOLOR := Buffer((p=4)?36:72,0) ; init dialog
    
    ; If (IsObject(custColorObj)) {
        ; Loop 16 {
            ; custColor := RGB_BGR(custColorObj[A_Index])
            ; NumPut "UInt", custColor, CUSTOM, (A_Index-1) * 4
        ; }
    ; }
    
    ; NumPut "UInt", CHOOSECOLOR.size, CHOOSECOLOR, 0             ; lStructSize
    ; NumPut "UPtr", hwnd,             CHOOSECOLOR, p             ; hwndOwner
    ; NumPut "UInt", RGB_BGR(color),   CHOOSECOLOR, 3 * p         ; rgbResult
    ; NumPut "UPtr", CUSTOM.ptr,       CHOOSECOLOR, 4 * p         ; lpCustColors
    ; NumPut "UInt", disp,             CHOOSECOLOR, 5 * p         ; Flags
    
    ; if !DllCall("comdlg32\ChooseColor", "UPtr", CHOOSECOLOR.ptr, "UInt")
        ; return -1
    
    ; custColorObj := []
    ; Loop 16 {
        ; newCustCol := NumGet(CUSTOM, (A_Index-1) * 4, "UInt")
        ; custColorObj.InsertAt(A_Index, RGB_BGR(newCustCol))
    ; }
    
    ; Color := NumGet(CHOOSECOLOR, 3 * A_PtrSize, "UInt")
    ; return Format("0x{:06X}",RGB_BGR(color))
    
    ; RGB_BGR(c) {
        ; return ((c & 0xFF) << 16 | c & 0xFF00 | c >> 16)
    ; }
; }
; */

FontSelect(fontObject:= "",hwnd:=0,Effects:=1) { ;Calls several windows intergrated DLLs to allow user to pick their font
		
			fontObject := (fontObject="") ? Map() : fontObject
			logfont := Buffer((A_PtrSize = 4) ? 60 : 92,0)
			uintVal := DllCall("GetDC","uint",0)
			LogPixels := DllCall("GetDeviceCaps","uint",uintVal,"uint",90)
			Effects := 0x041 + (Effects ? 0x100 : 0)
			
			fntName := fontObject.Has("name") ? fontObject["name"] : ""
			fontBold := fontObject.Has("bold") ? fontObject["bold"] : 0
			fontBold := fontBold ? 700 : 400
			fontItalic := fontObject.Has("italic") ? fontObject["italic"] : 0
			fontUnderline := fontObject.Has("underline") ? fontObject["underline"] : 0
			fontStrikeout := fontObject.Has("strike") ? fontObject["strike"] : 0
			fontSize := fontObject.Has("size") ? fontObject["size"] : 10
			fontSize := fontSize ? Floor(fontSize*LogPixels/72) : 16
			c := fontObject.Has("color") ? fontObject["color"] : 0
			
			c1 := Format("0x{:02X}",(c&255)<<16)	; convert RGB colors to BGR for input
			c2 := Format("0x{:02X}",c&65280)
			c3 := Format("0x{:02X}",c>>16)
			fontColor := Format("0x{:06X}",c1|c2|c3)
			
			NumPut "uint", fontSize, logfont
			NumPut "uint", fontBold, "char", fontItalic, "char", fontUnderline, "char", fontStrikeout, logfont, 16
			
			choosefont := Buffer(A_PtrSize=8?104:60,0), cap := choosefont.size
			NumPut "UPtr", hwnd, choosefont, A_PtrSize
			offset1 := (A_PtrSize = 8) ? 24 : 12
			offset2 := (A_PtrSize = 8) ? 36 : 20
			offset3 := (A_PtrSize = 4) ? 6 * A_PtrSize : 5 * A_PtrSize
			
			fontArray := Array([cap,0,"Uint"],[logfont.ptr,offset1,"UPtr"],[effects,offset2,"Uint"],[fontColor,offset3,"Uint"])
			
			for index,value in fontArray
				NumPut value[3], value[1], choosefont, value[2]
			
			if (A_PtrSize=8) {
				strput(fntName,logfont.ptr+28,"UTF-16")
				r := DllCall("comdlg32\ChooseFont","UPtr",CHOOSEFONT.ptr) ; cdecl 
				fntName := strget(logfont.ptr+28,"UTF-16")
			} else {
				strput(fntName,logfont.ptr+28,32,"UTF-8")
				r := DllCall("comdlg32\ChooseFontA","UPtr",CHOOSEFONT.ptr) ; cdecl
				fntName := strget(logfont.ptr+28,32,"UTF-8")
			}
			
			if !r
				return false
			
			fontObj := Map("bold",16,"italic",20,"underline",21,"strike",22)
			for a,b in fontObj
				fontObject[a] := NumGet(logfont,b,"UChar")
			
			fontObject["bold"] := (fontObject["bold"] < 188) ? 0 : 1
			
			; c  := NumGet(choosefont,(A_PtrSize=4)?6*A_PtrSize:5*A_PtrSize,"UInt") ; convert from BGR to RBG for output
			; c1 := Format("0x{:02X}",(c&255)<<16), c2 := Format("0x{:02X}",c&65280), c3 := Format("0x{:02X}",c>>16)
			; c  := Format("0x{:06X}",c1|c2|c3)
			; fontObject["color"] := c
			
			fontSize := NumGet(choosefont,A_PtrSize=8?32:16,"UInt") / 10 ; 32:16
			fontObject["size"] := fontSize
			fontHwnd := DllCall("CreateFontIndirect","uptr",logfont.ptr) ; last param "cdecl"
			fontObject["name"] := fntName
			
			logfont := "", choosefont := ""
			
			If (!fontHwnd) {
				fontObject := ""
				return 0
			} 
			
			Else {
				fontObject["hwnd"] := fontHwnd
				b := fontObject["bold"] ? "bold" : ""
				i := fontObject["italic"] ? "italic" : ""
				s := fontObject["strike"] ? "strike" : ""
				;c := fontObject["color"] ? "c" fontObject["color"] : ""
				z := fontObject["size"] ? "s" fontObject["size"] : ""
				u := fontObject["underline"] ? "underline" : ""
				fullStr := b "|" i "|" s "|" c "|" z "|" u
				str := ""
				Loop Parse fullStr, "|"
					If (A_LoopField)
						str .= A_LoopField " "
				fontObject["str"] := "norm " Trim(str)
				
				return fontObject

			}
		}


latcheck()
SetTimer latcheck, 2000

latcheck(){
Try{
if WinExist("Latitude - Account Work Form"){
Global LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
} ;if
Else{
try{
While !WinExist("Latitude - Account Work Form"){
try{
sleep 500
if WinExist("Latitude - Account Work Form"){
Global LatitudeEl := UIA.ElementFromHandle("Latitude - Account Work Form")
      } ;if
     } ;try
    } ;while
   } ;try
  } ;Else
 } ;try
} ;latcheck