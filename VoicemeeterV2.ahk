/*
    VOICEMEETER REMOTE API WRAPPER (AutoHotkey v2)
	@version : 1.2
	@author : Rian
    -----------------------------------------------------------------------
    USAGE EXAMPLES:
    vm  := Voicemeeter()            ; Initialize (Singleton)
    vm1 := Voicemeeter() 			; Same object as vm object
    [ STRIPS & BUSES ]
    vm.strip[0].Mute := 1           ; Set Mute
    vm.strip[0].Solo += 1           ; Toggle Solo (Smart Toggle)
    vm.strip[1].Gain := -10.5       ; Set Gain (dB)
    vm.bus[0].Label := "Headset"    ; Set Label (String)
    
    [ SPECIAL COMMANDS ]
    vm.ShowOrHide()                 ; Toggle GUI visibility
    vm.RestartEngine()              ; Restart Audio Engine
    vm.command.Button[0].State := 1 ; Set Macro Button State (untested)
    vm.command.Save := "C:\settings.xml" ; Save Configuration
    
    [ PROCESS INFO ]
    MsgBox vm.exe                   ; Get current executable name
    MsgBox vm.pid                   ; Get Voicemeeter Process ID
    -----------------------------------------------------------------------
*/
class Voicemeeter {
	static _instance := 0
	static File_Dir => this._GetVoicemeeterDir()
    static DLL_PATH => (A_PtrSize = 8) 
        ? this.File_Dir "\VoicemeeterRemote64.dll" 
        : this.File_Dir "\VoicemeeterRemote.dll"
    
	_type		:= 0
	_lastType	:= 0
	logged_in	:= false
	inputs		:= 0
	outputs		:= 0
	hDLL		:= 0
	fn			:= Map()
	
	; Return the same object every time (singleton)
	static Call(*) {
		if !this._instance
			this._instance := super.Call()
		return this._instance
	}
    
    type {
        get {
            ; Index: 0=None, 1=Basic(3in/2out), 2=Banana(5in/5out), 3=Potato(8in/8out)
            static input_map  := [0, 3, 5, 8]
            static output_map := [0, 2, 5, 8]

            rawType := 0
            res := DllCall(this.fn["VBVMR_GetVoicemeeterType"], "Int*", &rawType, "Int")

            if (res >= 0 && rawType > 0)
                this._lastType := rawType
            this._type := rawType

            ; Guard: unknown type — reset counts and return early
            if (rawType < 1 || rawType > 3) {
                this.inputs  := 0
                this.outputs := 0
                return this._type
            } 

            this.inputs  := input_map[this._type + 1]
            this.outputs := output_map[this._type + 1]

            return this._type
        }
    }
    
	typeName => Map(0,"None", 1,"Basic", 2,"Banana", 3,"Potato").Get(this.type, "Unknown")
	
	__New() {
        this.connected := false
		
        if !this._LoadDLL()
            throw Error("Failed to load Voicemeeter DLL. Check path: " Voicemeeter.DLL_PATH)

        if (this._Login() < 0)
            throw Error("Voicemeeter login failed.")
		
		_ := this.type
		
        ; Root nodes
        this.strip    := VMNode(this, "Strip")
        this.bus      := VMNode(this, "Bus")
        this.fx       := VMNode(this, "Fx")
        this.patch    := VMNode(this, "Patch")
        this.option   := VMNode(this, "Option")
        this.recorder := VMNode(this, "Recorder")
        this.vban     := VMNode(this, "vban")
        this.command  := VMNode(this, "Command")
		

		this._proc := VoicemeeterProcess(this)
    }
	
	__Get(Name, _) => HasProp(this._proc, Name) ? this._proc.%Name% : ""
	
	__Call(Name, Params) => HasMethod(this._proc, Name) ? this._proc.%Name%(Params*) : ""

    __Delete() {
        if this.hDLL {
            this._Logout()
            DllCall("FreeLibrary", "Ptr", this.hDLL)
            this.hDLL := 0
        }
		; BUG 2 FIX: Break the circular reference Voicemeeter → _proc → api → Voicemeeter
		; before clearing _instance. AHK v2 uses reference counting — without this,
		; the cycle keeps both objects alive and __Delete is never called by the GC.
		if this.HasOwnProp("_proc")
			this._proc.api := 0
        Voicemeeter._instance := 0
    }
	
	_LoadDLL() {
        if !(this.hDLL := DllCall("LoadLibrary", "Str", Voicemeeter.DLL_PATH, "Ptr"))
            return false

        rawFuncs := "
        (
            VBVMR_Login,VBVMR_Logout,VBVMR_RunVoicemeeter,VBVMR_GetVoicemeeterType,
            VBVMR_GetVoicemeeterVersion,VBVMR_IsParametersDirty,VBVMR_GetParameterFloat,
            VBVMR_SetParameterFloat,VBVMR_GetParameterStringA,VBVMR_GetParameterStringW,
            VBVMR_SetParameterStringA,VBVMR_SetParameterStringW,VBVMR_SetParameters,
            VBVMR_SetParametersW,VBVMR_GetLevel,VBVMR_GetMidiMessage,VBVMR_SendMidiMessage,
            VBVMR_MacroButton_IsDirty,VBVMR_MacroButton_GetStatus,VBVMR_MacroButton_SetStatus,
            VBVMR_Output_GetDeviceNumber,VBVMR_Output_GetDeviceDescA,VBVMR_Output_GetDeviceDescW,
            VBVMR_Input_GetDeviceNumber,VBVMR_Input_GetDeviceDescA,VBVMR_Input_GetDeviceDescW,
            VBVMR_AudioCallbackRegister,VBVMR_AudioCallbackStart,VBVMR_AudioCallbackStop,
            VBVMR_AudioCallbackUnregister
        )"

        cleanFuncs := RegExReplace(rawFuncs, "\s+", "")

        for funcName in StrSplit(cleanFuncs, ",") {
            if (funcName == "")
                continue

            ptr := DllCall("GetProcAddress", "Ptr", this.hDLL, "AStr", funcName, "Ptr")
            if !ptr
                throw Error("Failed to resolve DLL export: " funcName)
            this.fn[funcName] := ptr
        }
        return true
    }

    static _GetVoicemeeterDir() {
		static cachedPath := ""
        if (cachedPath != "")
            return cachedPath
		rootKeys := [
			"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
			"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
		]

		for root in rootKeys {
			Loop Reg, root, "K" {
				if InStr(A_LoopRegName, "VB:Voicemeeter") {
					try {
						rawPath := StrReplace(RegRead(root "\" A_LoopRegName, "UninstallString"), '"')
						if (rawPath != "") {
							SplitPath(rawPath, , &dir)
							if DirExist(dir)
								return cachedPath := dir
						}
					}
				}
			}
		}

		return cachedPath := (EnvGet("ProgramFiles(x86)") || A_ProgramFiles) "\VB\Voicemeeter"
	}

	_Login() {
		this.logged_in := true
		return DllCall(this.fn["VBVMR_Login"], "Int")
	}

	_Logout() {
		this.logged_in := false
		return DllCall(this.fn["VBVMR_Logout"], "Int")
	}

    EnsureConnected() {
        if (this.type > 0) {
            this.connected := true
            if (!this._proc.exe)             ; ← lazy: only detect when not yet found
                this._proc._DetectExeFromDLL()
            return true
        }
        this.connected := false
        return false
    }
		
	WaitForNotDirty(maxMs := 500) {
		start := A_TickCount
		Loop {
			if (DllCall(this.fn["VBVMR_IsParametersDirty"], "Int") == 0)
				return true
			if (A_TickCount - start >= maxMs)
				return false
			Sleep 10
		}
	}
	
	GetFloat(p) {
		if (DllCall(this.fn["VBVMR_IsParametersDirty"], "Int"))
			this.WaitForNotDirty()

		v := 0.0
		DllCall(this.fn["VBVMR_GetParameterFloat"], "AStr", p, "Float*", &v, "Int")
		return v
	}

    SetFloat(p, v) => DllCall(this.fn["VBVMR_SetParameterFloat"], "AStr", p, "Float", Float(v), "Int")
    
    SetString(p, v) => DllCall(this.fn["VBVMR_SetParameterStringW"], "AStr", p, "WStr", String(v), "Int")

	GetString(p) {
		if (DllCall(this.fn["VBVMR_IsParametersDirty"], "Int"))
			this.WaitForNotDirty()
        buf := Buffer(1024, 0)
        DllCall(this.fn["VBVMR_GetParameterStringW"], "AStr", p, "Ptr", buf.Ptr, "Int")
        return StrGet(buf, "UTF-16")
    }
}

;==================================================================
; Voicemeeter Nodes
;==================================================================
class VMNode {
    __New(vm, prefix) {
        this.DefineProp("_vm", {value: vm})
        this.DefineProp("_prefix", {value: prefix})
        this.DefineProp("_cache", {value: Map()})
    }

    __Item[idx] {
        get {
            if !this._cache.Has(idx)
                this._cache[idx] := VMNode(this._vm, this._prefix "[" idx "]")
            return this._cache[idx]
        }
    }

    __Get(name, params) {
        part := name
        for p in params
            part .= "[" p "]"
        
        fullPath := this._prefix (this._prefix ? "." : "") part

        ; Sub-namespaces: properties that return a child VMNode for further dot-chaining.
        ; Based on the API parameter list — these names always precede a further ".Property".
        ; e.g. Strip[i].Comp.Threshold, Strip[i].Pitch.On, Bus[i].Device.WDM
        static containers := "i)^(Comp|Gate|Denoiser|Pitch|GainLayer|EQ|Device|App|Mode|Patch|Color|Fx|outstream|instream|Buffer|Button|Preset|DialogShow|ArmStrip|ArmBus)$"
        if (params.Length > 0 || name ~= containers)
            return VMNode(this._vm, fullPath)
			
        ; String-typed parameters (readable via GetParameterStringW).
        ; Device.name / Device.sr are reached through the Device container above.
        ; Write-only strings (Save, Load, FadeTo, FadeBy) are handled in __Set only.
        static stringParams := "i)^(Label|name|ip|sr|channel|bit)$"
        if (name ~= stringParams)
            return this._vm.GetString(fullPath)

        ; FLOAT DEFAULT
        return this._vm.GetFloat(fullPath)
    }

    __Set(name, params, val) {
        ; Build full path including any bracket params (e.g. GainLayer[2], ArmStrip[0])
        part := name
        for p in params
            part .= "[" p "]"
        fullPath := this._prefix (this._prefix ? "." : "") part

        ; --- Array-value string commands: FadeTo/FadeBy/AppGain/AppMute ---
        ; e.g. strip[0].FadeTo := [-10, 500]  →  SetString("Strip[0].FadeTo", "(-10, 500)")
        if (name ~= "i)^(FadeTo|FadeBy|AppGain|AppMute)$") && (Type(val) = "Array")
            return this._vm.SetString(fullPath, "(" val[1] ", " val[2] ")")

        ; --- Write-only string params (Command paths and others) ---
        ; These accept a string value directly (file paths, stream names, IP addresses, labels).
        static writeStrings := "i)^(Save|Load|SaveBUSEQ|LoadBUSEQ|SaveStripEQ|LoadStripEQ|Label|name|ip|FileNameAttr|load|goto|WDM|KS|MME|ASIO)$"
        if (name ~= writeStrings) || (!IsNumber(val) && Type(val) = "String")
            return this._vm.SetString(fullPath, val)

        ; --- Boolean toggles: clamp out-of-range values (e.g. from += 1) to 0/1 ---
        ; Covers all 0-or-1 parameters from the API doc across Strip, Bus, Command, VBAN, Recorder.
        static toggles := "i)^(Mute|Solo|Mono|MC|A[1-5]|B[1-3]|On|PostReverb|PostDelay|PostFx1|PostFx2|Sel|Monitor|EQ\.On|Lock|Eject|Reset|Show|Shutdown|Record|Play|Stop|Loop|MakeUp|State|StateOnly|Trigger|Recall|Restart)$"
        if (name ~= toggles) {
            current := this._vm.GetFloat(fullPath)
            ; Smart toggle: if val is out of 0/1 range (e.g. from +=1), flip current state
            if (val > 1 || val < 0)
                val := (current = 1 ? 0 : 1)
        }

        return this._vm.SetFloat(fullPath, val)
    }

    __Cast(target) {
        if (target = "Number" || target = "Float" || target = "Integer")
            return this._vm.GetFloat(this._prefix)
        return this._vm.GetString(this._prefix)
    }
}

; ==============================================================================
; PROCESS & WINDOW INTERACTION CLASS
; ==============================================================================

class VoicemeeterProcess {
	; BUG 1 FIX: Removed the singleton pattern (static _instance + static Call).
	;
	; The old pattern had two problems:
	;   a) On re-init after Voicemeeter.__Delete, VoicemeeterProcess(newVM) would
	;      return the stale _instance whose .api still pointed to the freed/unloaded
	;      old Voicemeeter — including its freed DLL handle. Any subsequent DLL call
	;      through _proc crashed into freed memory.
	;   b) super.Call(apiInstance) passes args to Object's allocator, which creates
	;      a plain Object, not a VoicemeeterProcess, so _instance was set to an
	;      Object with none of the expected properties or methods.
	;
	; VoicemeeterProcess doesn't need its own singleton: it is always constructed
	; inside Voicemeeter.__New, and Voicemeeter is itself the singleton. _proc is
	; therefore effectively unique through Voicemeeter's own guarantee.

	exe  := ""
    pid  := 0
    hwnd := 0
	
	__New(apiInstance) {
        if (apiInstance == "")
            apiInstance := Voicemeeter()
            
        this.api := apiInstance
        this._DetectExeFromDLL()
    }

    _DetectExeFromDLL() {
        if (this.api.type == 0) {          ; ← check type directly, no circular call
            return false
        }
        this.api.connected := true

        old := DetectHiddenWindows(true)
        class:= "VBCABLE0Voicemeeter0MainWindow0"

        if (hWin := WinExist("ahk_class " class)) {
            this.hwnd := hWin
            this.pid := WinGetPID(hWin)
            this.exe := WinGetProcessName(hWin)
            DetectHiddenWindows(old)
            return true
        }

		; BUG 5 FIX: Use _lastType (cached) instead of calling the type property getter.
		; The type getter does a full DLL call and has the side effect of updating
		; this.api.inputs and this.api.outputs — undesirable inside a detection helper.
		; _lastType is set whenever a valid type was last seen, which is sufficient here.
        static names := ["Voicemeeter", "VoicemeeterPro", "Voicemeeter8"]
        vType := this.api._lastType
		if (vType < 1 || vType > 3) {
			DetectHiddenWindows(old)
			return false
		}
		base := names[vType]

        for suffix in ["", "_x64", "x64"] {
            target := base suffix ".exe"
            if (pid := ProcessExist(target)) {
                this.pid := pid
                this.exe := target
                DetectHiddenWindows(old)
                return true
            }
        }

        DetectHiddenWindows(old)
        return false
    }

    GetOrSetVBAN(newState := "") {
        if (newState = "")
            return this.api.GetFloat("vban.Enable")
        if (this.api.SetFloat("vban.Enable", newState) == 0)
			return Float(newState)
		return this.api.GetFloat("vban.Enable")
    }

    RestartEngine(*) => this.api.SetFloat("Command.Restart", 1)

    RestartVoicemeeter(*) {
        this.Shutdown()
		if (this.exe) {
			try ProcessWaitClose(this.exe, 3)
		}
        this.hwnd := 0
        this.pid := 0
        return this.ShowOrHide()
    }
	
	Shutdown(*) => this.api.SetFloat("Command.Shutdown", 1)
	
    ShowOrHide(*) {
        if !this.exe
            return false
        old := DetectHiddenWindows(true)

        ; Case 1: We have a known hwnd and it still exists — show or hide it
        if (this.hwnd && WinExist("ahk_id " this.hwnd)) {
            if WinActive("ahk_id " this.hwnd)
                WinHide("ahk_id " this.hwnd)
            else {
                WinShow("ahk_id " this.hwnd)
                WinActivate("ahk_id " this.hwnd)
            }
            DetectHiddenWindows(old)
            return true
        }

        ; Case 2: Window exists by exe name but we lost the hwnd — refresh and recurse once
        if (hwnd := WinExist("ahk_exe " this.exe)) {
            this.hwnd := hwnd
            this.pid := WinGetPID(hwnd)
            DetectHiddenWindows(old)
            return this.ShowOrHide()
        }

        ; Case 3: Window not found — try to launch via DLL first, then fall back to Run
		vType := this.api._lastType
		if (vType > 0) {
            DllCall(this.api.fn["VBVMR_RunVoicemeeter"], "Int", vType, "Int")
            
            if (this.hwnd := WinWait("ahk_exe " this.exe, , 5)) {
                this.pid := WinGetPID(this.hwnd)
                WinShow("ahk_id " this.hwnd)
                WinActivate("ahk_id " this.hwnd)
                DetectHiddenWindows(old)
                return true
            }
			; BUG 4 FIX: Return here if DLL launch attempt timed out.
			; Previously the code fell through to Run(path), which would launch
			; a second Voicemeeter instance on top of the one RunVoicemeeter already
			; started (it may still be initialising when WinWait gives up at 5s).
            DetectHiddenWindows(old)
            return false
        }

        ; Case 4: No DLL launch available — run the exe directly
        path := Voicemeeter.File_Dir "\" this.exe
        try {
			Tooltip "Try"
            Run(path, , , &newpid)
            this.pid := newpid
            if (this.hwnd := WinWait("ahk_exe " this.exe, , 5)) {
                WinShow("ahk_id " this.hwnd)
                WinActivate("ahk_id " this.hwnd)
                DetectHiddenWindows(old)
                return true
            }
        }
        DetectHiddenWindows(old)
        return false
    }
}
