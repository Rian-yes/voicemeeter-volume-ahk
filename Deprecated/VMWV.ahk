#Requires AutoHotkey v2.1-alpha.1+
#SingleInstance Force
SetWorkingDir(A_ScriptDir)
DetectHiddenWindows(true)
Persistent true
#Include %A_ScriptDir%\Audio.ahk

OnExit(OnScriptExit)

OnMessage(0x0218, VM_RestartVoicemeeterEngine)
OnMessage(0x404, OnDeviceTrayClick) ;WM_TRAY_NOTIFY
OnMessage(0x0219, AnyDeviceChange)
SetTimer Poller,5000

; Global Object to hold VM state
global VM := {
    hDLL: 0,
    dllPath: "",
    type: 0,
    pid: 0,
    exe: "",
    hwnd: 0,
    inputs: 0,
    outputs: 0,
	VBAN : 0,
}

global DEV := {
	anydevicescanner : 0,
	restartonlaunch : 0,
	limit_gain: 0,
    syncmute: 1,
    linear_Vol: 1,
    restartAudioOnDevice: 1,
    rememberVol: 0,
	allDevices : [],
	audioDevices : [],
	DeviceMap : Map(),
	CheckedStates : Map(),
	CurrentGains : Map()
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

;Script Functions
SaveConfig(*) {
    global DEV
    iniPath := A_ScriptDir "\settings.ini"
    
    ; 1. Save Simple Toggle Settings
    settings := ["anydevicescanner", "restartonlaunch", "limit_gain", "syncmute", "linear_Vol", "restartAudioOnDevice", "rememberVol"]
    for key in settings
        IniWrite(DEV.%key%, iniPath, "Settings", key)

    ; 2. Save Persistent Maps
    ; We skip alldevices/audioDevices as they should be fresh every run
    for mapName in ["DeviceMap", "CheckedStates", "CurrentGains"] {
        try IniDelete(iniPath, mapName)
        for key, value in DEV.%mapName%
            IniWrite(value == "" ? "NULL" : value, iniPath, mapName, key)
    }
}

LoadConfig() {
    global DEV
    iniPath := A_ScriptDir "\settings.ini"
    if !FileExist(iniPath)
        return

    ; 1. Load Settings
    settings := ["anydevicescanner", "restartonlaunch", "limit_gain", "syncmute", "linear_Vol", "restartAudioOnDevice", "rememberVol"]
    for key in settings {
        try val := IniRead(iniPath, "Settings", key)
        if IsSet(val)
            DEV.%key% := Integer(val)
    }

    ; 2. Load Maps
    for mapName in ["CheckedStates", "CurrentGains"] {
        try {
            section := IniRead(iniPath, mapName)
            loop parse section, "`n", "`r" {
                if !(pos := InStr(A_LoopField, "=")) 
                    continue
                
                key := SubStr(A_LoopField, 1, pos - 1)
                val := SubStr(A_LoopField, pos + 1)
                
                ; Convert back to actual numbers for logic/math
                DEV.%mapName%[key] := IsNumber(val) ? Number(val) : (val == "NULL" ? "" : val)
            }
        }
    }
	OnScriptInit()
	return 1
}

;Start

if (!LoadConfig()) {
    SetTimer(Reinitial, -4000)
}

;Stop
OnScriptExit(*) {
    SaveConfig()
    VM_Logout()
}
OnScriptInit(*) {
	if (!VM_Init()) {
		SetTimer Reinitial,-2000
	}
	InitializeVolumeSync()
	CheckMenus()
	if (DEV.restartonlaunch) {
		VM_RestartVoicemeeterEngine()
    }
	if (DEV.rememberVol) {
		SyncVoicemeeterToWindows()
    }
}


;---------- WM Messages ---------------
AnyDeviceChange(wParam, lParam, msg, hwnd) {
    SetTimer(CheckDevices, -20)
}

OnDeviceTrayClick(wParam, lParam, msg, hwnd) {
    if (lParam = 0x0205) {
        CoordMode "Mouse", "Screen"
        MouseGetPos(&Misc.MouseX, &Misc.MouseY)
    }
}

; --- Initialization Logic ---

Poller() {
	VM.type:=VM_GetVoicemeeterVersion()
	static input_map  := [3, 5, 8] ; Basic, Banana, Potato
    static output_map := [2, 5, 8]
    
    if (VM.type==0 || VM.inputs != input_map[VM.type] || VM.outputs != output_map[VM.type] )
		SetTimer Reinitial,-2000

}
Generate_menu(){
	; 1. Create Submenus first
	if !DEV.HasProp("Menus")
        DEV.Menus := {}
	DEV.Menus.BindVolume := Menu()
    DEV.Menus.RestartAudio := Menu()
    DEV.Menus.Settings := Menu()

	; --- Submenu: Bind Windows Volume to ---
	DEV.Menus.BindVolume.Add("INPUTS", Dummy)
	DEV.Menus.BindVolume.Disable("INPUTS")
	DEV.Menus.BindVolume.Add() ; Separator
	DEV.Menus.BindVolume.Add("OUTPUTS", Dummy)
	DEV.Menus.BindVolume.Disable("OUTPUTS")
	DEV.Menus.BindVolume.Add()

	; --- Submenu: Restart Audio Engine on ---
	DEV.Menus.RestartAudio.Add("Any Device Change", RestartAudioOnAction)
	DEV.Menus.RestartAudio.Add("Audio Device Change", RestartAudioOnAction)
	DEV.Menus.RestartAudio.Check("Audio Device Change")
	DEV.Menus.RestartAudio.Add("Resume from Standby", RestartAudioOnAction)
	DEV.Menus.RestartAudio.Add("App Launch", RestartAudioOnAction)

	; --- Submenu: Settings ---
	DEV.Menus.Settings.Add("MAIN SETTINGS", Dummy)
	DEV.Menus.Settings.Disable("MAIN SETTINGS")
	DEV.Menus.Settings.Add()
	DEV.Menus.Settings.Add("Automatically Start with Windows", ToggleSetting)
	DEV.Menus.Settings.Add("Limit Gain to 0dB", ToggleSetting)
	DEV.Menus.Settings.Add("Use Linear Volume Scaling", ToggleSetting)
	DEV.Menus.Settings.Check("Use Linear Volume Scaling")
	DEV.Menus.Settings.Add("Sync Mute", ToggleSetting)
	DEV.Menus.Settings.Add("FIXES", Dummy)
	DEV.Menus.Settings.Disable("FIXES")
	DEV.Menus.Settings.Add()
	DEV.Menus.Settings.Add("Restore Volume on Launch", ToggleSetting)
	DEV.Menus.Settings.Add(" ", Dummy)
	DEV.Menus.Settings.Disable(" ")

	; --- Main Tray Menu Setup ---
	TMenu := A_TrayMenu
	TMenu.Delete() 

	TMenu.Add("Voicemeeter Windows Volume", Dummy)
	TMenu.Disable("Voicemeeter Windows Volume")

	TMenu.Add("Bind Windows Volume to", DEV.Menus.BindVolume)
	TMenu.Add("Restart Audio Engine on", DEV.Menus.RestartAudio)
	TMenu.Add("Settings", DEV.Menus.Settings)

	TMenu.Add() ; Separator

	TMenu.Add("Voicemeeter Functions", Dummy)
	TMenu.Disable("Voicemeeter Functions")

	TMenu.Add("Restart Voicemeeter", VM_RestartVoicemeeter)
	TMenu.Add("Show Voicemeeter", VM_ShowOrHideVoicemeeter)
	TMenu.Add("Restart Audio Engine", VM_RestartVoicemeeterEngine)
	TMenu.Add("VBAN State", VM_MenuToggleVBAN)

	TMenu.Add()
	TMenu.Add("Exit", (*) => ExitApp()) ; Concise arrow function for Exit

	TMenu.Default := "Show Voicemeeter"
	A_Icontip := "Voicemeeter Volume Control"
}

CheckMenus() {
    global DEV, VM, Misc
    
    ; 1. Sync Strip Checkmarks
    for displayName, stripAddress in DEV.DeviceMap {
        try {
            if (DEV.CheckedStates.Has(stripAddress) && DEV.CheckedStates[stripAddress])
                DEV.Menus.BindVolume.Check(displayName)
            else
                DEV.Menus.BindVolume.Uncheck(displayName)
        }
    }

    ; 2. Settings Menu Sync
    sMenu := DEV.Menus.Settings
    try {
        (DEV.limit_gain)  ? sMenu.Check("Limit Gain to 0dB") : sMenu.Uncheck("Limit Gain to 0dB")
        (DEV.syncmute)    ? sMenu.Check("Sync Mute") : sMenu.Uncheck("Sync Mute")
        (DEV.rememberVol) ? sMenu.Check("Restore Volume on Launch") : sMenu.Uncheck("Restore Volume on Launch")
        (DEV.linear_Vol)  ? sMenu.Check("Use Linear Volume Scaling") : sMenu.Uncheck("Use Linear Volume Scaling")
    }

    ; 3. Auto-start Logic Validation
    shortcutValid := 0
    if FileExist(Misc.shortcutPath) {
        try {
            FileGetShortcut(Misc.shortcutPath, &target)
            shortcutValid := (target == A_ScriptFullPath)
        }
    }
    try sMenu.Check("Automatically Start with Windows", shortcutValid ? 1 : 0)
    DEV.autostart := shortcutValid

    ; 4. Restart Conditions
    rMenu := DEV.Menus.RestartAudio
    try {
        if (DEV.anydevicescanner) {
            rMenu.Check("Any Device Change")
            rMenu.Uncheck("Audio Device Change")
        } else {
            rMenu.Uncheck("Any Device Change")
            rMenu.Check("Audio Device Change")
        }
        (DEV.restartonlaunch) ? rMenu.Check("App Launch") : rMenu.Uncheck("App Launch")
    }
	if (VM_GetOrSetVBAN()) {
	    A_TrayMenu.Check("VBAN State")
	}
}
Reinitial() {
    static waiter := 0
    static attempts := 0
    voicemeeterType := VM_GetVoicemeeterVersion()
    
    bindMenu := DEV.Menus.BindVolume
    if (voicemeeterType != 0) {
        if (attempts < 3) { 
            attempts++
            ToolTip("VoiceMeeter detected... initializing engine (" attempts ")")
            SetTimer(Reinitial, -1500) ; Check again in 1.5s
            return
        }

        
        SetTimer(Reinitial, 0) 
        waiter := 0
        attempts := 0
		static input_map  := [3, 5, 8] ; Basic, Banana, Potato
		static output_map := [2, 5, 8]
		
		VM.inputs  := input_map[VM.type]
		VM.outputs := output_map[VM.type]
        ToolTip("VoiceMeeter Loaded and Menu Ready")
        SetTimer(() => ToolTip(), -3000)		
		ReGenerateDevices()
        
    } else {
        ; Voicemeeter is not open or DLL isn't connecting yet
        if (!waiter) {
            waiter := 1
			bindMenu.Delete()
            
            ; 2. Add a disabled placeholder so the menu isn't empty
            bindMenu.Add("(Waiting for Voicemeeter...)", Dummy)
            bindMenu.Disable("(Waiting for Voicemeeter...)")
            
            ; 3. Clear the DeviceMap so no stale logic runs
            DEV.DeviceMap := Map()
        }
        
        attempts := 0
        ; Continue the loop: Check again in 2 seconds
        SetTimer(Reinitial, -2000) 
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
            currentList := ConvertToArray(GetDevicesFromEnum(2).List)
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

        ; 3. Compare Lists
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
            
            ToolTip(msg)
            SetTimer(() => ToolTip(), -5000) ; Increased to 5s for readability

            ; 5. Trigger Restart Logic
            ; Restart only if we are in "Audio" mode OR if "Any" mode is specifically asked to restart
            if (DEV.restartAudioOnDevice || DEV.anydevicescanner  ) {
                VM_RestartVoicemeeterEngine()
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

; Helper: Clean string to array
ConvertToArray(str) {
    arr := []
    Loop Parse, str, "`n", "`r" {
        if (Trim(A_LoopField) != "")
            arr.Push(Trim(A_LoopField))
    }
    return arr
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

    ; 1. Get the state and the address (e.g., "Bus[0]")
    isChecked := MyMenu.ToggleCheck(ItemName)
    targetAddress := DEV.DeviceMap[ItemName]
    
    ; 2. Link Logic: Find every menu item that shares this address
    ; This ensures A1 and A2 toggle together in Basic
    for menuLabel, address in DEV.DeviceMap {
        if (address == targetAddress) {
            DEV.CheckedStates[address] := isChecked ; Update the Map
            try {
				(isChecked) ? MyMenu.Check(menuLabel) : MyMenu.UnCheck(menuLabel) 
			}
        }
    }
    
    SaveConfig() 
    SyncVoicemeeterToWindows()
    
    CoordMode "Menu", "Screen"
	SetTimer(MaintainMenuZOrder, 20)
    A_TrayMenu.Show(Misc.MouseX, Misc.MouseY,0)
}

;Hmm AHK v2.1 alpha "FIX"
MaintainMenuZOrder() {
    ; Find ALL active menu windows (including submenus)
    hwnds := WinGetList("ahk_class #32768")
    
    if (hwnds.Length == 0) {
        SetTimer(MaintainMenuZOrder, 0) ; Stop timer 
        return
    }

    for hwnd in hwnds {
        ; Check if it's already topmost to avoid flickering (0x8 is WS_EX_TOPMOST)
        exStyle := WinGetExStyle(hwnd)
        if !(exStyle & 0x8) {
            ; SWP_NOSIZE(1) | SWP_NOMOVE(2) | SWP_NOACTIVATE(16) = 0x13
            DllCall("SetWindowPos", "Ptr", hwnd, "Ptr", -1, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", 0x13)
        }
    }
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
    if (!A_IsCompiled) {
        MsgBox("Please compile the script before using the 'Start with Windows' option.")
        return
    }

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
        
    ; Optional: Trigger an INI save here if you have one
    ; SaveSettings() 
}
 
 
 
 ; -----------Menu Functions--------------------
 ; Add Voicemeeter Inputs to Menu
AddVoicemeeterInputs() {
    VM_WaitForNotDirty()
    hwInputs := VM.inputs-VM_GetVoicemeeterVersion()
    
    Loop VM.inputs {
        stripIndex := A_Index - 1
        stripName  := VM_GetString("Strip[" stripIndex "].device.name")
        stripLabel := VM_GetLabel("Strip[" stripIndex "]")
        
        if (A_Index > hwInputs) {
            vIndex := A_Index - hwInputs
            vDefaultName := (vIndex == 1) ? "Voicemeeter VAIO Input" 
                          : (vIndex == 2) ? "Voicemeeter AUX Input" 
                          : "Voicemeeter VAIO3 Input"
            
            mainTitle := (stripLabel != "") ? stripLabel : "Virtual Input " vIndex
            deviceTitle := vDefaultName
        } 
        else {
            mainTitle := (stripLabel != "") ? stripLabel : "Hardware Input " . A_Index
            deviceTitle := (stripName != "") ? stripName : "No Device"
        }

        menuLabel := mainTitle . " : <" . deviceTitle . ">"
        
        DEV.Menus.BindVolume.Insert("OUTPUTS", menuLabel, BindVolumeAction)
        DEV.DeviceMap[menuLabel] := "Strip[" stripIndex "]"
    }
}
; Add Voicemeeter Outputs to Menu
AddVoicemeeterOutputs() {
    VM_WaitForNotDirty()
    if (VM_GetVoicemeeterVersion() == 1) {
        ; --- A1 Setup ---
        busNameA1 := VM_GetString("Bus[0].device.name")
        busLabelA := VM_GetLabel("Bus[0]")
        mainA1 := (busLabelA != "") ? busLabelA : "Hardware Output A1"
        deviceA1 := (busNameA1 != "") ? busNameA1 : "Internal Master Clock"
        
        labelA1 := mainA1 . " : <" . deviceA1 . "> [shared]"
        DEV.Menus.BindVolume.Add(labelA1, BindVolumeAction)
        DEV.DeviceMap[labelA1] := "Bus[0]" ; Use the full label as the key

        ; --- A2 Setup ---
        busNameA2 := VM_GetString("Bus[1].device.name")
        if (busNameA2 != "") {
            mainA2 := "Hardware Output A2" 
            labelA2 := mainA2 . " : <" . busNameA2 . "> [shared]"
            DEV.Menus.BindVolume.Add(labelA2, BindVolumeAction)
            DEV.DeviceMap[labelA2] := "Bus[0]" ; Shared with A1
        }
        
        ; --- B Setup ---
        busLabelB := VM_GetLabel("Bus[2]") ; Note: Basic uses Bus[2] internally for Virtual B
        mainB := (busLabelB != "") ? busLabelB : "Virtual Output B"
        labelB := mainB . " : <Voicemeeter Output>"
        
        DEV.Menus.BindVolume.Add(labelB, BindVolumeAction)
        DEV.DeviceMap[labelB] := "Bus[1]"
    } else {
        hwOutputs := VM.outputs-VM_GetVoicemeeterVersion()
        
        Loop VM.outputs {
            busIndex := A_Index - 1
            busName  := VM_GetString("Bus[" busIndex "].device.name")
            busLabel := VM_GetLabel("Bus[" busIndex "]")
            
            if (A_Index > hwOutputs) {
                vIndex := A_Index - hwOutputs
                vInternalName := (vIndex == 1) ? "Voicemeeter Output" 
                               : (vIndex == 2) ? "Voicemeeter Aux Output" 
                               : "Voicemeeter VAIO3 Output"
                
                mainTitle := (busLabel != "") ? busLabel : "Virtual Output B" . vIndex
                deviceTitle := vInternalName
            } else {
                mainTitle := (busLabel != "") ? busLabel : "Hardware Output A" . A_Index
                if (busIndex == 0 && busName == "")
                    deviceTitle := "Internal Master Clock"
                else
                    deviceTitle := (busName != "") ? busName : "No Device"
            }

            menuLabel := mainTitle . " : <" . deviceTitle . ">"
            DEV.Menus.BindVolume.Add(menuLabel, BindVolumeAction)
            DEV.DeviceMap[menuLabel] := "Bus[" busIndex "]"
        }
    }
}
 
ReGenerateDevices() {
	global DEV
	bindMenu := DEV.Menus.BindVolume
	bindMenu.Delete()
	
	bindMenu.Add("INPUTS", Dummy)
	bindMenu.Disable("INPUTS")
	bindMenu.Add() ; Separator
	
	bindMenu.Add("OUTPUTS", Dummy)
	bindMenu.Disable("OUTPUTS")
	bindMenu.Add()
	DEV.DeviceMap := Map()
	AddVoicemeeterInputs() 
	AddVoicemeeterOutputs()
	CheckMenus()
} 
 
 
 ; --- Functions ---
VM_Init() {
	Generate_menu()
    if (!VM_LoadDLL())
        return false
    
    if (VM_Login())
        return false
        
    if (!VM_GetVoicemeeterVersion())    
        return false
    
    if (!VM_DetectExeFromDLL())    
        return false
		
    static input_map  := [3, 5, 8] ; Basic, Banana, Potato
    static output_map := [2, 5, 8]
    
    VM.inputs  := input_map[VM.type]
    VM.outputs := output_map[VM.type]
	
	if (DEV.restartonlaunch)
		VM_RestartVoicemeeterEngine()
		
	ReGenerateDevices()
    return true
}
VM_MenuToggleVBAN(ItemName, ItemPos, MyMenu) {
    ; Get current state from the DLL (or the cached VM object)
    currentState := VM_GetFloat("vban.Enable")
    
    ; Determine the new state (flip it)
    newState := (currentState == 0.0) ? 1.0 : 0.0
    
    ; Apply to Voicemeeter
    VM_GetOrSetVBAN(newState)
    
    ; Update the Menu UI
    if (newState == 1.0)
        MyMenu.Check(ItemName)
    else
        MyMenu.Uncheck(ItemName)
        
}

VM_LoadDLL() {
    if (VM.hDLL)
        return true
    
    VM.dllPath := A_Is64bitOS 
        ? "C:\Program Files (x86)\VB\Voicemeeter\VoicemeeterRemote64.dll" 
        : "C:\Program Files (x86)\VB\Voicemeeter\VoicemeeterRemote.dll"
    
    VM.hDLL := DllCall("LoadLibrary", "Str", VM.dllPath, "Ptr")
    
    if (!VM.hDLL)
        MsgBox("Error: Failed to load DLL! Check path: " VM.dllPath)
    
    return VM.hDLL != 0
}

VM_Login() {
    return DllCall(VM.dllPath "\VBVMR_Login", "Int")
}

VM_Logout(*) {
    if (VM.hDLL)
        DllCall(VM.dllPath "\VBVMR_Logout", "Int")
}

VM_GetVoicemeeterVersion() {
    vmtype := 0
    DllCall(VM.dllPath "\VBVMR_GetVoicemeeterType", "Int*", &vmtype, "Int")
    VM.type := vmtype
    return VM.type
}

VM_DetectExeFromDLL() {
    if (!VM.type)
        return false
        
    ; 1. Try Window Class first
    classes := ["VBCABLE0Voicemeeter0MainWindow0", "VMR_MainForm", "VBC_MainForm", "VBC8_MainForm"]
    for className in classes {
        if (hWin := WinExist("ahk_class " className)) {
            VM.pid := WinGetPID(hWin)
            VM.exe := WinGetProcessName(hWin)
            return true
        }
    }
    
    ; 2. Fallback to name variations
    static names := ["Voicemeeter", "VoicemeeterPro", "Voicemeeter8"]
    base := names[VM.type]
    
    for suffix in ["", "_x64", "x64"] {
        target := base . suffix . ".exe"
        if (pid := ProcessExist(target)) {
            VM.pid := pid
            VM.exe := WinGetProcessName("ahk_pid " pid)
            return true
        }
    }
    return false
}


VM_GetOrSetVBAN(newState:="") {
	VM_WaitForNotDirty()
	if (newState==""){
		VM.VBAN:=VM_GetFloat("vban.Enable")
		return VM.VBAN
	}
    VM_SetFloat("vban.Enable", VM.VBAN)
	return VM.VBAN
}

VM_RestartVoicemeeterEngine(*) {
    return VM_SetFloat("Command.Restart", 1)
}

VM_RestartVoicemeeter(*) {
    ; Close by HWND if it exists, otherwise fallback to PID
    if (VM.pid && ProcessExist(VM.pid)) {
        ProcessClose(VM.pid)
    }
    
    ProcessWaitClose(VM.pid || "Voicemeeter.exe", 5) 
    
    ToolTip("Restarting Voicemeeter...")
    VM.hwnd := 0  ; Reset the handle
    VM.pid := 0 
    VM_ShowOrHideVoicemeeter()
}

VM_ShowOrHideVoicemeeter(*) {
    if !(VM.exe)
        return

    DetectHiddenWindows(true)
    
	vmClass := "ahk_class VBCABLE0Voicemeeter0MainWindow0"
	
    ; 1. Check if our stored HWND is still valid
    if (VM.hwnd && WinExist(VM.hwnd)) {
        if WinActive(VM.hwnd) {
            WinHide(VM.hwnd)
            if WinExist("ahk_exe explorer.exe")
                WinActivate("ahk_exe explorer.exe")
        } else {
            WinShow(VM.hwnd)
            WinActivate(VM.hwnd)
        }
    } 
    ; 2. If HWND is invalid, check if the window exists by EXE
	else if (hwnd := WinExist(vmClass . " ahk_exe " VM.exe)) {
        VM.hwnd := hwnd
        VM.pid := WinGetPID(hwnd)
        VM_ShowOrHideVoicemeeter() 
    }
    ; 3. Launch if nothing exists
    else {
        path := "C:\Program Files (x86)\VB\Voicemeeter\" . VM.exe
        try {
            Run(path, , , &newpid)
            VM.pid := newpid
            
            ; Wait for the window and capture its NEW HWND
            if (VM.hwnd := WinWait("ahk_exe " VM.exe, , 5)) {
                WinShow(VM.hwnd)
                WinActivate(VM.hwnd)
                ToolTip("Voicemeeter Launched")
            }
        } catch {
            MsgBox("Error: Could not find Voicemeeter at " . path)
            return
        }
        SetTimer(Reinitial, -2000) 
    }
    SetTimer(() => ToolTip(), -3000)
}
; -------------- API Commands ----------------------

VM_SetGain(strip, level) => VM_SetFloat(strip ".Gain", level)
VM_GetGain(strip)       => VM_GetFloat(strip ".Gain")
VM_SetMute(strip, state) => VM_SetFloat(strip ".Mute", state)
VM_GetMute(strip)       => VM_GetFloat(strip ".Mute")
VM_GetLabel(strip)      => VM_GetString(strip ".Label")

VM_GetFloat(paramName) {
    VM_WaitForNotDirty()
    val := 0.0 
	DllCall(VM.dllPath "\VBVMR_GetParameterFloat", "AStr", paramName, "Float*", &val, "Int")
    return val
}

VM_SetFloat(paramName, floatVal) {
    return DllCall(VM.dllPath "\VBVMR_SetParameterFloat", "AStr", paramName, "Float", floatVal)
}

VM_GetString(paramName) {
    VM_WaitForNotDirty()
    strBuf := Buffer(1024) 
    DllCall(VM.dllPath "\VBVMR_GetParameterStringW", "AStr", paramName, "Ptr", strBuf, "Int")
    return StrGet(strBuf, "UTF-16")
}

VM_SetString(paramName, strVal) {
    return DllCall(VM.dllPath "\VBVMR_SetParameterStringW", "AStr", paramName, "WStr", strVal)
}

VM_WaitForNotDirty() {
    while DllCall(VM.dllPath "\VBVMR_IsParametersDirty") {
        Sleep(20)
		if (A_Index > 30) ; Safety timeout (200ms)
            break
    }
}

;-----------Additional Functions -----------------

GetDevicesFromEnum(device_type := 0) {
    ; device_type: playback = 0, capture = 1, all = 2
    Devices := Map()
    List := ""
    
    ; CLSID_MMDeviceEnumerator and IID_IMMDeviceEnumerator
    IMMDeviceEnumerator := ComObject("{BCDE0395-E52F-467C-8E3D-C4579291692E}", "{A95664D2-9614-4F35-A746-DE8DB63617E6}")
    
    ; IMMDeviceEnumerator::EnumAudioEndpoints (Index 3)
    ; 0x1 = DEVICE_STATE_ACTIVE
    ComCall(3, IMMDeviceEnumerator, "UInt", device_type, "UInt", 0x1, "PtrP", &IMMDeviceCollection := 0)
    
    ; IMMDeviceCollection::GetCount (Index 3)
    ComCall(3, IMMDeviceCollection, "UIntP", &Count := 0)
    
    Loop Count {
        ; IMMDeviceCollection::Item (Index 4)
        ComCall(4, IMMDeviceCollection, "UInt", A_Index - 1, "PtrP", &IMMDevice := 0)

        ; IMMDevice::GetId (Index 5)
        ComCall(5, IMMDevice, "PtrP", &pBuffer := 0)
        DeviceID := StrGet(pBuffer, "UTF-16")
        DllCall("Ole32.dll\CoTaskMemFree", "Ptr", pBuffer)

        ; IMMDevice::OpenPropertyStore (Index 4) - 0x0 = STGM_READ
        ComCall(4, IMMDevice, "UInt", 0x0, "PtrP", &IPropertyStore := 0)
        ObjRelease(IMMDevice)
    
        ; IPropertyStore::GetValue (Index 5)
        ; PKEY_Device_FriendlyName = {a45c254e-df1c-4efd-8020-67d146a850e0} 14
        PROPVARIANT := Buffer(A_PtrSize == 8 ? 24 : 16, 0)
        PROPERTYKEY := Buffer(20, 0)
        DllCall("Ole32.dll\CLSIDFromString", "Str", "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", "Ptr", PROPERTYKEY)
        NumPut("UInt", 14, PROPERTYKEY, 16)
        
        ComCall(5, IPropertyStore, "Ptr", PROPERTYKEY, "Ptr", PROPVARIANT)
        
        ; Get the PWSTR from the PROPVARIANT structure (offset 8)
        pwszVal := NumGet(PROPVARIANT, 8, "Ptr")
        if (pwszVal) {
            DeviceName := StrGet(pwszVal, "UTF-16")
            DllCall("Ole32.dll\CoTaskMemFree", "Ptr", pwszVal)
        } else {
            DeviceName := "Unknown Device"
        }
        
        ObjRelease(IPropertyStore)
        Devices[DeviceName] := DeviceID
        List .= DeviceName "`n"
    }
    
    ObjRelease(IMMDeviceCollection)
    return {List: List, Map: Devices}
}

AdjustCheckedVolumes(Offset) {
	RefreshGainCache()
	for stripAddress, isChecked in DEV.CheckedStates {
        if (isChecked) {
			newGain := DEV.CurrentGains[stripAddress] + Offset
            clampedGain := Min(Max(newGain, -60), 12)
			DEV.CurrentGains[stripAddress] := clampedGain
            VM_SetGain(stripAddress, clampedGain)
        }
    }
	static RefreshGainCache() {
    ; Only refresh if the menu ISN'T open to avoid interfering with your scrolls
        for _,stripAddress in DEV.DeviceMap {
            DEV.CurrentGains[stripAddress] := VM_GetFloat(stripAddress ".Gain")
        }
	}
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
        for addr, checked in DEV.CheckedStates {
            if (checked) {
                VM_SetGain(addr, db)
                if (DEV.syncmute)
					VM_SetMute(addr, Notified.bMuted)
            }
        }
    }
}

InitializeVolumeSync() {
    try {
        AUDIO_RESOURCES.enumerator := IMMDeviceEnumerator()
        AUDIO_RESOURCES.device     := AUDIO_RESOURCES.enumerator.GetDefaultAudioEndpoint(0, 0)
        AUDIO_RESOURCES.volume     := AUDIO_RESOURCES.device.Activate(IAudioEndpointVolume)
        AUDIO_RESOURCES.sink       := VoicemeeterVolumeSync()
        
        ; Registering the persistent sink
        AUDIO_RESOURCES.volume.RegisterControlChangeNotify(AUDIO_RESOURCES.sink)
    } catch Error as e {
        MsgBox("Hook Registration FAILED:`n`n" e.Message)
    }
}

SyncVoicemeeterToWindows() {
    try {
        ; Get current Windows values (0-100 and 0/1)
        TEMP:= {
			fMasterVolume:SoundGetVolume()/ 100,
			bMuted :SoundGetMute()
			}
        
		VoicemeeterVolumeSync.ExecuteSync(TEMP)
        
        ToolTip("Voicemeeter Synced to Windows: " Round(TEMP.fMasterVolume*100) "%")
    } catch Error as e {
        ToolTip("Sync Failed: " e.Message)
    }
    SetTimer(() => ToolTip(), -2000)
}
