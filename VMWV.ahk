#Requires AutoHotkey v2.1-alpha.1+
#SingleInstance Force
SetWorkingDir(A_ScriptDir)
DetectHiddenWindows(true)
Persistent true
#Include %A_ScriptDir%\Audio.ahk
#Include %A_ScriptDir%\VoicemeeterV2.ahk


global DEV := {
	anydevicescanner : 0,
	restartonlaunch : 0,
	limit_gain: 0,
	syncmute: 1,
	linear_Vol: 1,
	restartAudioOnDevice: 1,
	rememberVol: 0,
	previousVol : -1,
	allDevices : [],
	audioDevices : [],
	DeviceMap : Map(),
	CheckedStates : Map(),
	ActiveSync : Map(),
	CurrentGains : Map(),
	Menus : {}
}

global AUDIO_RESOURCES := {
	enumerator: "",
	device: "",
	volume: "",
	sink: ""
}

global Misc := {
	MouseX: -1,
	MouseY: -1,
	shortcutPath : A_Startup "\" RegExReplace(A_ScriptName, "\.exe$", "") ".lnk",
	shortcutValid : 0
}

;Start

OnExit(OnScriptExit)
OnMessage(0x0219, AnyDeviceChange)
OnScriptInit()

;============================================
;   Script Functions
;============================================
SaveConfig(*) {
	global DEV
	iniPath := A_ScriptDir "\settings.ini"
	if FileExist(iniPath)
		FileDelete iniPath
		
	if (DEV.rememberVol) 
		DEV.previousVol := SoundGetVolume()
	
	settings := ["anydevicescanner", "restartonlaunch", "limit_gain", "syncmute", "linear_Vol", "restartAudioOnDevice", "rememberVol", "previousVol"]
	for key in settings
		IniWrite(DEV.%key%, iniPath, "Settings", key)

	for mapName in ["CheckedStates", "CurrentGains"] {
		try IniDelete(iniPath, mapName)
		for key, value in DEV.%mapName% {
			; Ensure we don't save an object by accident
			if IsObject(value) 
				continue 
			IniWrite(value == "" ? "NULL" : value, iniPath, mapName, key)
		}
	}
}

LoadConfig() {
	global DEV
	iniPath := A_ScriptDir "\settings.ini"
	if !FileExist(iniPath)
		return

	; Load Simple Settings
	settings := ["anydevicescanner", "restartonlaunch", "limit_gain", "syncmute", "linear_Vol", "restartAudioOnDevice", "rememberVol", "previousVol"]
	for key in settings {
		try val := IniRead(iniPath, "Settings", key, "")
		if (val !== "")
			DEV.%key% := Integer(val)
	}

	; Load Maps
	for mapName in ["CheckedStates", "CurrentGains"] {
		try {
			section := IniRead(iniPath, mapName)
			Loop Parse, section, "`n", "`r" {
				if !(pos := InStr(A_LoopField, "=")) 
					continue
					
				key := SubStr(A_LoopField, 1, pos - 1)
				val := SubStr(A_LoopField, pos + 1)
					
				DEV.%mapName%[key] := IsNumber(val) ? Number(val) : (val == "NULL" ? "" : val)
			}
		}
	}
	return 1
}

;Stop
OnScriptExit(*) {
	try VM.__Delete()
	SaveConfig()
}

OnScriptInit(*) {
	LoadConfig()
	if (!VM_Init())
		Reinitial()
	Generate_menu()
	ReGenerateDevices()
	InitializeVolumeSync()
	CheckMenus()
	CheckDevices()
	if (DEV.rememberVol || DEV.previousVol != -1) {
		SoundSetVolume(DEV.previousVol)
		; Force a sync 
		SyncVoicemeeterToWindows()
	}
	
	;Modeless Menu Hook Fix
	static hHook := DllCall("SetWinEventHook"
		, "UInt", 0x8002, "UInt", 0x8002 ; Range of events (only show)
		, "Ptr", 0, "Ptr", CallbackCreate(OnMenuShow, "F")
		, "UInt", 0, "UInt", 0, "UInt", 0)

	; The event-driven function
	OnMenuShow(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime) {
		; Only act if the object shown is a Window and its class is a Menu
		if (idObject = 0 && WinGetClass(hwnd) = "#32768") {
			; Set TopMost + NoActivate (0x13 = NOSIZE | NOMOVE | NOACTIVATE)
			DllCall("SetWindowPos", "Ptr", hwnd, "Ptr", -1, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x13)
		}
	}
	SetTimer Poller,5000
	return 1
}

Poller() {
	if (VM.type==0) {
		SetTimer Poller,0
		Reinitial()	
	}
}

Reinitial() {
	static waiter := 0
	voicemeeterType := VM.type
	
	bindMenu := DEV.Menus.BindVolume
	TMenu := A_TrayMenu
	if (voicemeeterType != 0) {
		SetTimer(Reinitial, 0) 
 		waiter := 0
		TMenu.Enable("Restart Voicemeeter")
		TMenu.Enable("Show Voicemeeter")
		TMenu.Enable("Restart Audio Engine")
		TMenu.Enable("VBAN State")
		TMenu.Enable("Shutdown Voicemeeter")
		SetTimer Poller, 5000		
		ReGenerateDevices()
		
	} else {
		; Voicemeeter is not open or DLL isn't connecting yet
		if (!waiter) {
			waiter := 1
			bindMenu.Delete()
			
			; Add a placeholder
			bindMenu.Add("(Waiting for Voicemeeter...)", Dummy)
			bindMenu.Disable("(Waiting for Voicemeeter...)")
			TMenu.Disable("VBAN State")
			TMenu.Disable("Restart Audio Engine")
			TMenu.Disable("Shutdown Voicemeeter")
			if (VM.exe){ 
				TMenu.Enable("Restart Voicemeeter")
				TMenu.Enable("Show Voicemeeter")
			}
			DEV.DeviceMap := Map()
		}
		
		; Continue the loop: Check again in 2 seconds
		SetTimer(Reinitial, -5000) 
	}
}


;---------- WM Messages ---------------
AnyDeviceChange(wParam, lParam, msg, hwnd) {
	SetTimer(CheckDevices, -20)
}

; --- Initialization Logic ---

Generate_menu(){
	; 1. Create Submenus first
	bindMenu	 := DEV.Menus.BindVolume   := Menu()
	restartMenu  := DEV.Menus.RestartAudio := Menu()
	settingsMenu := DEV.Menus.Settings	 := Menu()

	; --- Submenu: Bind Windows Volume to ---
	bindMenu.Add("INPUTS", Dummy)
	bindMenu.Disable("INPUTS")
	bindMenu.Add() ; Separator
	bindMenu.Add("OUTPUTS", Dummy)
	bindMenu.Disable("OUTPUTS")
	bindMenu.Add()

	; --- Submenu: Restart Audio Engine on ---
	restartMenu.Add("Any Device Change", RestartAudioOnAction)
	restartMenu.Add("Audio Device Change", RestartAudioOnAction)
	restartMenu.Check("Audio Device Change")
	restartMenu.Add("Resume from Standby", RestartAudioOnAction)
	restartMenu.Disable("Resume from Standby")
	restartMenu.Add("App Launch", RestartAudioOnAction)

	; --- Submenu: Settings ---
	settingsMenu.Add("MAIN SETTINGS", Dummy)
	settingsMenu.Disable("MAIN SETTINGS")
	settingsMenu.Add()
	settingsMenu.Add("Automatically Start with Windows", ToggleSetting)
	settingsMenu.Add("Limit Gain to 0dB", ToggleSetting)
	settingsMenu.Add("Use Linear Volume Scaling", ToggleSetting)
	settingsMenu.Check("Use Linear Volume Scaling")
	settingsMenu.Add("Sync Mute", ToggleSetting)
	settingsMenu.Add("FIXES", Dummy)
	settingsMenu.Disable("FIXES")
	settingsMenu.Add()
	settingsMenu.Add("Restore Volume on Launch", ToggleSetting)
	settingsMenu.Add(" ", Dummy)
	settingsMenu.Disable(" ")

	; --- Main Tray Menu Setup ---
	TMenu := A_TrayMenu
	TMenu.Delete() 

	TMenu.Add("Voicemeeter Windows Volume", Dummy)
	TMenu.Disable("Voicemeeter Windows Volume")
	TMenu.Add()

	TMenu.Add("Bind Windows Volume to", bindMenu)
	TMenu.Add("Restart Audio Engine on", restartMenu)
	TMenu.Add("Settings", settingsMenu)


	TMenu.Add("Voicemeeter Functions", Dummy)
	TMenu.Disable("Voicemeeter Functions")
	TMenu.Add()

	TMenu.Add("Restart Voicemeeter", ObjBindMethod(VM, "RestartVoicemeeter"))
	TMenu.Disable("Restart Voicemeeter")
	TMenu.Add("Show Voicemeeter",	ObjBindMethod(VM, "ShowOrHide"))
	TMenu.Disable("Show Voicemeeter")
	TMenu.Add("Restart Audio Engine", ObjBindMethod(VM, "RestartEngine"))
	TMenu.Disable("Restart Audio Engine")
	TMenu.Add("Shutdown Voicemeeter", (*) => (VM.Shutdown(), Poller()))
	TMenu.Disable("Shutdown Voicemeeter")
	TMenu.Add("VBAN State", VM_MenuToggleVBAN)
	TMenu.Disable("VBAN State")

	TMenu.Add()
	TMenu.Add("Exit", (*) => ExitApp()) ; Concise arrow function for Exit

	TMenu.Default := "Show Voicemeeter"
	A_Icontip := "Voicemeeter Volume Control"
}

CheckMenus() {
	global DEV, VM, Misc
	
	DEV.ActiveSync := Map()
	bindMenu := DEV.Menus.BindVolume
	; 1. Sync Strip Checkmarks
	for displayName, nodeObj in DEV.DeviceMap {
	try {
		key := nodeObj._prefix
		state := (DEV.CheckedStates.Has(key) && DEV.CheckedStates[key])
			
		if (state) {
			bindMenu.Check(displayName)
				
			; --- NEW: Populate ActiveSync ---
			; We pre-calculate the strings here so ExecuteSync stays ultra-fast
			DEV.ActiveSync[key] := {
				gainStr: key ".Gain",
				muteStr: key ".Mute"
			}
		} else {
			bindMenu.Uncheck(displayName)
		}
	}
}

	; 2. Settings Menu Sync
	settingsMenu := DEV.Menus.Settings
	try {
		(DEV.limit_gain)  ? settingsMenu.Check("Limit Gain to 0dB") : settingsMenu.Uncheck("Limit Gain to 0dB")
		(DEV.syncmute)	? settingsMenu.Check("Sync Mute") : settingsMenu.Uncheck("Sync Mute")
		(DEV.rememberVol) ? settingsMenu.Check("Restore Volume on Launch") : settingsMenu.Uncheck("Restore Volume on Launch")
		(DEV.linear_Vol)  ? settingsMenu.Check("Use Linear Volume Scaling") : settingsMenu.Uncheck("Use Linear Volume Scaling")
	}

	; 3. Auto-start Logic Validation
	shortcutValid := 0
	if FileExist(Misc.shortcutPath) {
		try {
			FileGetShortcut(Misc.shortcutPath, &target)
			shortcutValid := (target == A_ScriptFullPath)
		}
	}
	try settingsMenu.Check("Automatically Start with Windows", shortcutValid ? 1 : 0)
	DEV.autostart := shortcutValid

	; 4. Restart Conditions
	restartMenu := DEV.Menus.RestartAudio
	try {
		if (DEV.anydevicescanner) {
			restartMenu.Check("Any Device Change")
			restartMenu.Uncheck("Audio Device Change")
		} else {
			restartMenu.Uncheck("Any Device Change")
			restartMenu.Check("Audio Device Change")
		}
		(DEV.restartonlaunch) ? restartMenu.Check("App Launch") : restartMenu.Uncheck("App Launch")
	}
	if (VM.GetOrSetVBAN()) {
		A_TrayMenu.Check("VBAN State")
	}
	if (VM.type) {
		TMenu := A_TrayMenu
		TMenu.Enable("Restart Voicemeeter")
		TMenu.Enable("Show Voicemeeter")
		TMenu.Enable("Restart Audio Engine")
		TMenu.Enable("VBAN State")
		TMenu.Enable("Shutdown Voicemeeter")
	}
}

CheckDevices(*) {
	static scanInitAll := true
	static scanInitAudio := true
	static isScanning := false

	if (isScanning)
		return
	isScanning := true

	try {
		local currentList := []
		local targetStorage := ""
		local isFirstRun := false

		; 1. Determine Scan Mode
		if (DEV.anydevicescanner) {
			currentList := GetAllSystemDevices()
			targetStorage := "allDevices"
			sleep 3000
			if (scanInitAll || DEV.allDevices.Length == 0) {
				isFirstRun := true
				scanInitAll := false
			}
		} else if (DEV.restartAudioOnDevice) {
			; Only scan audio if that toggle is actually ON
			currentList := GetAudioDeviceNames(2) ;2 = eAll
			targetStorage := "audioDevices"
			if (scanInitAudio || DEV.audioDevices.Length == 0) {
				isFirstRun := true
				scanInitAudio := false
			}
		} else {
			return ; Both toggles off
		}

		; 2. Initialize baseline
		if (isFirstRun) {
			DEV.%targetStorage% := currentList
			return
		}

		added := []
		removed := []
		oldList := DEV.%targetStorage%
		; Optimization: Use a temporary Map for O(1) lookups if the list is huge
		for item in currentList {
			if !HasValue(oldList, item)
				added.Push(item)
		}

		for item in oldList {
			if !HasValue(currentList, item)
				removed.Push(item)
		}
		; 4. Show ToolTips (Moved outside the VM restart check)
		if (added.Length > 0 || removed.Length > 0) {
			msg := ""
			if (added.Length > 0) {
				msg .= "Connected:`n"
				for item in added
					msg .= "+ " item "`n"
			}
			if (removed.Length > 0) {
				msg .= (msg ? "`n" : "") . "Disconnected:`n"
				for item in removed
					msg .= "- " item "`n"
			}
			
			; ToolTip(msg)
			; SetTimer(() => ToolTip(), -5000)

			; 5. Trigger Restart Logic
			; Restart only if we are in "Audio" mode OR if "Any" mode is specifically asked to restart
			if (DEV.restartAudioOnDevice || DEV.anydevicescanner  ) {
				VM.RestartEngine()
			}
		}

		; 6. Update Storage
		DEV.%targetStorage% := currentList

	} finally {
		isScanning := false
	}
}

GetAllSystemDevices() {
	deviceList := []

	try {
		wmi := ComObjGet("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
		if !IsObject(wmi) {
			wmi := ComObjGet("winmgmts:")
			if !IsObject(wmi) {
				return []
			}
		}

		query := "SELECT Name FROM Win32_PnPEntity"
		colItems := wmi.ExecQuery(query)
		if !IsObject(colItems) {
			return []
		}

		enum := colItems._NewEnum()
		count := 0
		while enum(&item) {
			try {
				name := item.Name
				if (name != "") {
					deviceList.Push(name)
					count++
				}
			} catch {
				continue ; Skip items that error out
			}
		}
	} catch Any as e {
		MsgBox("CRITICAL ERROR:`nType: " . type(e) . "`nMessage: " . e.Message . "`nLine: " . e.Line)
		return []
	}

	return deviceList
}


HasValue(haystack, needle) {
	for index, value in haystack {
		if (value == needle)
			return index
	}
	return 0
}

;---------------------Menu Handlers -----------------------------
Dummy(*){
return
}

BindVolumeAction(ItemName, ItemPos, MyMenu) {
	Critical
	global DEV, Misc
	
	if !DEV.DeviceMap.Has(ItemName)
		return

	targetNode := DEV.DeviceMap[ItemName]
	
	isChecked := MyMenu.ToggleCheck(ItemName)
	
	; link Logic: Find every menu item that shares this address
	; This ensures A1 and A2 toggle together in Basic
	for menuLabel, nodeObj in DEV.DeviceMap {
		if (nodeObj == targetNode) {
			DEV.CheckedStates[nodeObj._prefix] := isChecked
			
			; --- EFFICENCY BOOST ---
			; Store the object and its pre-built gain/mute strings 
			if (isChecked) {
				DEV.ActiveSync[nodeObj._prefix] := {
					gainStr: nodeObj._prefix ".Gain",
					muteStr: nodeObj._prefix ".Mute"
				}
			}
			else {
				DEV.ActiveSync.Delete(nodeObj._prefix)
			}
			try {
				(isChecked) ? MyMenu.Check(menuLabel) : MyMenu.UnCheck(menuLabel) 
			}
		}
	}

	SaveConfig() 
	SyncVoicemeeterToWindows()
	
	CoordMode "Menu", "Screen"
	A_TrayMenu.Show(Misc.MouseX, Misc.MouseY,0)
}

RestartAudioOnAction(itemName, itemPos, menuObj) {

	if (itemName == "App Launch") {
		DEV.restartonlaunch := !DEV.restartonlaunch
		(DEV.restartonlaunch) ? menuObj.Check(itemName) : menuObj.Uncheck(itemName)
			
	} else if (itemName == "Any Device Change") {
		DEV.anydevicescanner := !DEV.anydevicescanner
		
		if (DEV.anydevicescanner) {
			DEV.restartAudioOnDevice := 0
			menuObj.Check("Any Device Change")
			menuObj.Uncheck("Audio Device Change")
		} else {
			menuObj.Uncheck("Any Device Change")
		}
		
		SetTimer(CheckDevices, -20) 

	} else if (itemName == "Audio Device Change") {
		DEV.restartAudioOnDevice := !DEV.restartAudioOnDevice
		
		if (DEV.restartAudioOnDevice) {
			DEV.anydevicescanner := 0
			menuObj.Check("Audio Device Change")
			menuObj.Uncheck("Any Device Change")
		} else {
			menuObj.Uncheck("Audio Device Change")
		}
		
		SetTimer(CheckDevices, -20)
	} else {
		DEV.restartonlaunch := !DEV.restartonlaunch
	}
}

ToggleSetting(itemName, itemPos, menuObj) {
	
	if (itemName == "Limit Gain to 0dB") {
		Toggleandcheck("limit_gain", itemName, menuObj)
	}
	else if (itemName == "Automatically Start with Windows") {
		;@Ahk2Exe-IgnoreBegin
		if (!A_IsCompiled) {
			MsgBox("Please compile the script before using the 'Start with Windows' option.")
			return
		}
		;@Ahk2Exe-IgnoreEnd

		; Check if shortcut exists AND if it points to THIS current file
		Misc.shortcutValid := false
		if FileExist(Misc.shortcutPath) {
			try {
				FileGetShortcut(Misc.shortcutPath, &targetPath)
				if (targetPath == A_ScriptFullPath)
					Misc.shortcutValid := true
			}
		}

		if (Misc.shortcutValid) {
			; If it's valid and we clicked it, the user wants to DISABLE it
			FileDelete(Misc.shortcutPath)
			menuObj.Uncheck(itemName)
			DEV.autostart := 0
		} else {
			if FileExist(Misc.shortcutPath)
				FileDelete(Misc.shortcutPath)
				
			FileCreateShortcut(A_ScriptFullPath, Misc.shortcutPath)
			menuObj.Check(itemName)
			DEV.autostart := 1
		}
	}
	else if (itemName == "Sync Mute") {
		Toggleandcheck("syncmute", itemName, menuObj)
	}
	else if (itemName == "Use Linear Volume Scaling") {
		Toggleandcheck("linear_Vol", itemName, menuObj)
	}
	else {
		; Fallback for "Restore Volume on Launch" / "RememberVol"
		Toggleandcheck("rememberVol", itemName, menuObj)
	}
	SyncVoicemeeterToWindows()
}


; Helper function to handle the actual toggling and menu state
Toggleandcheck(propName, itemName, menuObj) {
	currentVal := DEV.%propName%
	newVal := !currentVal
	DEV.%propName% := newVal
	
	if (newVal)
		menuObj.Check(itemName)
	else
		menuObj.Uncheck(itemName)

	SaveConfig() 
}
 
 
 
; -----------Menu Functions--------------------
; Add Voicemeeter Inputs to Menu
AddVoicemeeterInputs() {
	VMtype:=VM.type
	if (VMtype == 0)
		return
		
	hwInputs := VM.inputs-VMtype
	
	Loop VM.inputs {
		i := A_Index - 1
		currentStrip := VM.strip[i]
		
		stripName  := currentStrip.device.name
		stripLabel := currentStrip.Label 
		
		if (A_Index > hwInputs) {
			vIndex := A_Index - hwInputs
			vDefaultName := (vIndex == 1) ? "Voicemeeter VAIO Input" 
				: (vIndex == 2) ? "Voicemeeter AUX Input" 
				: "Voicemeeter VAIO3 Input"
			
		mainTitle := (stripLabel != "") ? stripLabel : "Virtual Input " vIndex
		deviceTitle := vDefaultName
		} else {
			; --- Hardware Input Logic ---
			mainTitle := (stripLabel != "") ? stripLabel : "Hardware Input " . A_Index
			deviceTitle := (stripName != "") ? stripName : "No Device"
		}
	menuLabel := mainTitle . " : <" . deviceTitle . ">"

	DEV.Menus.BindVolume.Insert("OUTPUTS", menuLabel, BindVolumeAction)
	DEV.DeviceMap[menuLabel] := currentStrip
	}
}

; Add Voicemeeter Outputs to Menu
AddVoicemeeterOutputs() {
	VMtype:=VM.type
	if (VMtype == 0)
		return
		
	hwOutputs := VM.outputs-VMtype
	
	Loop VM.outputs {
		i := A_Index - 1
		bus := VM.bus[i]
		
		busName  := bus.device.name
		busLabel := bus.Label
		
		if (A_Index <= hwOutputs) {
			; --- Hardware Outputs (A) ---
			mainTitle := (busLabel != "") ? busLabel : "Hardware Output A" A_Index
			
			; Special case for A1 Master Clock
			if (i == 0 && busName == "")
				deviceTitle := "Internal Master Clock"
			else
				deviceTitle := (busName != "") ? busName : "No Device"
			
			; Basic Version logic: A1 and A2 share the same gain/bus[0] in the API
			; but show as separate devices in the UI.
			if (VMtype == 1 && A_Index == 2) {
				menuLabel := mainTitle " : <" busName "> [shared]"
				targetObj := VM.bus[0] ; Map A2 to Bus[0] logic
			} else {
				menuLabel := mainTitle " : <" deviceTitle ">"
				targetObj := bus
			}
			} else {
			; --- Virtual Outputs (B) ---
			vIndex := A_Index - hwOutputs
			vInternalName := (vIndex == 1) ? "Voicemeeter Output" 
				: (vIndex == 2) ? "Voicemeeter Aux Output" 
				: "Voicemeeter VAIO3 Output"
			
			mainTitle := (busLabel != "") ? busLabel : "Virtual Output B" vIndex
			deviceTitle := vInternalName
			menuLabel := mainTitle " : <" deviceTitle ">"
			targetObj := bus
		}
		DEV.Menus.BindVolume.Add(menuLabel, BindVolumeAction)
		DEV.DeviceMap[menuLabel] := targetObj
	}
}
 
ReGenerateDevices() {
	global DEV
	bindMenu := DEV.Menus.BindVolume
	if (VM.type != 0) {
		bindMenu.Delete()
		
		bindMenu.Add("INPUTS", Dummy)
		bindMenu.Disable("INPUTS")
		bindMenu.Add()
		
		bindMenu.Add("OUTPUTS", Dummy)
		bindMenu.Disable("OUTPUTS")
		bindMenu.Add()
		
		DEV.DeviceMap := Map()
		AddVoicemeeterInputs() 
		AddVoicemeeterOutputs()
		CheckMenus()
	} else {
		bindMenu.Delete()
		bindMenu.Add("(Waiting for Voicemeeter...)", Dummy)
		bindMenu.Disable("(Waiting for Voicemeeter...)")
	}	
} 
 
 
; --- Voicemeeter Functions ---

VM_Init() {
	global VM := Voicemeeter()
	OnMessage(0x0218, ObjBindMethod(VM, "RestartVoicemeeter"))
	
	if (VM.type || DEV.restartonlaunch)
		VM.RestartEngine()
	
	return VM.type
}

VM_MenuToggleVBAN(ItemName, ItemPos, MyMenu) {
	newState := (VM.GetOrSetVBAN() == 0.0) ? 1.0 : 0.0
	VM.GetOrSetVBAN(newState)

	if (newState == 1.0)
		MyMenu.Check(ItemName)
	else
		MyMenu.Uncheck(ItemName)
}

;-----------Additional Functions -----------------
;Audio.ahk version Enumeration of devices
;-------------------------------------------------
GetAudioDevices(dataFlow := 0, stateMask := 1) {
	deviceList := []
	enumerator := IMMDeviceEnumerator()
	
	try {
		; EnumAudioEndpoints returns an IMMDeviceCollection
		collection := enumerator.EnumAudioEndpoints(dataFlow, stateMask)
		
		; Loop through the collection (using the class's __Enum logic)
		for device in collection {
			deviceList.Push({
			Name: device.GetName(),
			ID:   device.GetId()
			})
		}
	} catch Error as e {
		OutputDebug("Audio Enum Error: " . e.Message)
	}
	return deviceList
}

GetAudioDeviceNames(dataFlow := 2) {
	names := []
	; We reuse your logic but only extract the Name string
	for device in GetAudioDevices(dataFlow) {
		names.Push(device.Name)
	}
	return names
}

; ----------CALLBACK-----------------
class VoicemeeterVolumeSync extends IAudioEndpointVolumeCallback {

	static vtable := [
		(this, iid, pobj) => !NumPut("ptr", this, pobj),
		(this) => 1,
		(this) => 1,
		["OnNotify", this.AUDIO_VOLUME_NOTIFICATION_DATA]
	]

	OnNotify(Notify) {
		data := {
		fMasterVolume: Notify.fMasterVolume,
		bMuted: Notify.bMuted
		}
		VoicemeeterVolumeSync.ExecuteSync(data)
		return 0
	}

	static ExecuteSync(Notified) {
		db := (DEV.linear_Vol) ? (Notified.fMasterVolume * (DEV.limit_gain ? 60 : 72)) - 60 : 20 * Log(((10**((DEV.limit_gain ? 0 : 12)/20)) - (10**(-60/20))) * Notified.fMasterVolume + (10**(-60/20)))
		for prefix, data in DEV.ActiveSync {
			vm.SetFloat(data.gainStr, db)
			if (DEV.syncmute)
				vm.SetFloat(data.muteStr, Notified.bMuted)
		}
	}
}

InitializeVolumeSync() {
	try {
		AUDIO_RESOURCES.enumerator	:= IMMDeviceEnumerator()
		AUDIO_RESOURCES.device		:= AUDIO_RESOURCES.enumerator.GetDefaultAudioEndpoint(0, 0)
		AUDIO_RESOURCES.volume		:= AUDIO_RESOURCES.device.Activate(IAudioEndpointVolume)
		AUDIO_RESOURCES.sink		:= VoicemeeterVolumeSync()
		
		; Registering the persistent sink
		AUDIO_RESOURCES.volume.RegisterControlChangeNotify(AUDIO_RESOURCES.sink)
	} catch Error as e {
		MsgBox("Hook Registration FAILED:`n`n" e.Message)
	}
}

SyncVoicemeeterToWindows() {
	try {
		TEMP:= {
			fMasterVolume:SoundGetVolume()/ 100,
			bMuted :SoundGetMute()
		}
		
		VoicemeeterVolumeSync.ExecuteSync(TEMP)
			
		; ToolTip("Voicemeeter Synced to Windows: " Round(TEMP.fMasterVolume*100) "%")
		; SetTimer(() => ToolTip(), -2000)
	} catch Error as e {
		MsgBox("Manual Sync to Voicemeeter Failed: " e.Message)
	} 
}