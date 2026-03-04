/*
    VOICEMEETER REMOTE API WRAPPER (AutoHotkey v2)
	@version : 1.0
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
    vm.command.Button[0].State := 1 ; Set Macro Button State
    vm.command.Save := "C:\cfg.xml" ; Save Configuration
    
    [ PROCESS INFO ]
    MsgBox vm.exe                   ; Get current executable name
    MsgBox vm.pid                   ; Get Voicemeeter Process ID
    -----------------------------------------------------------------------
*/
class Voicemeeter {
	static _instance := 0
    static DLL_PATH := (A_PtrSize = 8) 
        ? "C:\Program Files (x86)\VB\Voicemeeter\VoicemeeterRemote64.dll" 
        : "C:\Program Files (x86)\VB\Voicemeeter\VoicemeeterRemote.dll"
    
	_type   := 0
	inputs  := 0
	outputs := 0
	hDLL    := 0
    fn      := Map()
	
	;Return Same Object
	static Call() {
        if (this._instance == 0)
            this._instance := super.Call() 
        return this._instance 
    }
	
	type {
		get {
			static input_map  := [0, 3, 5, 8] ; Basic, Banana, Potato
			static output_map := [0, 2, 5, 8]
			rawType := 0
			res := DllCall(this.fn["VBVMR_GetVoicemeeterType"], "Int*", &rawType, "Int")
			this._type := (res < 0) ? 0 : rawType
			this.inputs  := input_map[this._type+1]
			this.outputs := output_map[this._type+1] 
			return this._type
		}
	}
    
	typeName {
        get {
            static typeMap := Map(0, "None", 1, "Basic", 2, "Banana", 3, "Potato")
            return typeMap.Has(this._type) ? typeMap[this._type] : "Unknown"
        }
    }
	
	__New() {
		if (Voicemeeter._instance != 0)
            return Voicemeeter._instance
		
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

    __Delete() {
        if this.hDLL {
            this._Logout()
            DllCall("FreeLibrary", "Ptr", this.hDLL)
        }
    }
	
	; Access VoicemeeterProcess Object methods
    exe  => (this._proc.exe || (this._proc._DetectExeFromDLL(), this._proc.exe))
    hwnd => (this._proc.hwnd || (this._proc._DetectExeFromDLL(), this._proc.hwnd))
    pid  => (this._proc.pid || (this._proc._DetectExeFromDLL(), this._proc.pid))
    
    ShowOrHide(*)         => this._proc.ShowOrHide()
    RestartVoicemeeter(*) => this._proc.RestartVoicemeeter()
    RestartEngine(*)      => this._proc.RestartEngine()
    GetOrSetVBAN(a:="",*) => this._proc.GetOrSetVBAN(a)
    Shutdown(*)  	   	  => this._proc.Shutdown()
    
	
	_LoadDLL() {
        if !(this.hDLL := DllCall("LoadLibrary", "Str", Voicemeeter.DLL_PATH, "Ptr"))
            return false

        for func in ["VBVMR_Login", "VBVMR_Logout", "VBVMR_GetVoicemeeterType", 
                     "VBVMR_GetParameterFloat", "VBVMR_SetParameterFloat", 
                     "VBVMR_GetParameterStringW", "VBVMR_SetParameterStringW", 
                     "VBVMR_IsParametersDirty"] {
            this.fn[func] := DllCall("GetProcAddress", "Ptr", this.hDLL, "AStr", func, "Ptr")
        }
        return true
    }
    
    _GetTypeName() {
        static typeMap := Map(0, "None", 1, "Basic", 2, "Banana", 3, "Potato", 6, "Potato x64")
        return typeMap.Has(this.type) ? typeMap[this.type] : "Unknown"
    }

    _Login()  => DllCall(this.fn["VBVMR_Login"], "Int")
    _Logout() => DllCall(this.fn["VBVMR_Logout"], "Int")

	EnsureConnected() {
        ; Use the property check directly
        if (this.type > 0) {
            this.connected := true
            return true
        }
        this.connected := false
        return false
    }
	
	;Core methods
	GetType() => this.type
		
    WaitForNotDirty() {
        Loop 30 {
            if (DllCall(this.fn["VBVMR_IsParametersDirty"], "Int") == 0)
                return true
            Sleep 10
        }
        return false
    }
	
    GetFloat(p) {
        this.WaitForNotDirty()
        v := Float(0)
        DllCall(this.fn["VBVMR_GetParameterFloat"], "AStr", p, "Float*", &v, "Int")
        return v
    }

    SetFloat(p, v) => DllCall(this.fn["VBVMR_SetParameterFloat"], "AStr", p, "Float", Float(v), "Int")
    
    SetString(p, v) => DllCall(this.fn["VBVMR_SetParameterStringW"], "AStr", p, "WStr", String(v), "Int")

    GetString(p) {
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
        if (name = "_vm" || name = "_prefix" || name = "_cache")
            return this.%name%

        part := name
        for p in params
            part .= "[" p "]"
        
        fullPath := this._prefix (this._prefix ? "." : "") part

        static containers := "i)^(outstream|instream|mode|app|item|patch|color|fx|recorder|vban|device|buffer|delay|sr|Slim|WDM|KS|MME|ASIO|Comp|Gate|Denoiser|Pitch|EQ|channel|cell|Reverb|Delay|Button|Preset|DialogShow)$"        
		if (params.Length > 0 || name ~= containers)
            return VMNode(this._vm, fullPath)
			
        static stringParams := "i)^(Label|FadeTo|FadeBy|AppGain|AppMute|goto|load|save|name|ip|LoadBUSEQ|SaveBUSEQ|LoadStripEQ|SaveStripEQ)$"        
		if (name ~= stringParams)
            return this._vm.GetString(fullPath)

        ; 3. FLOAT DEFAULT
        return this._vm.GetFloat(fullPath)
    }

    __Set(name, params, val) {
        fullPath := this._prefix (this._prefix ? "." : "") name

        if (name ~= "i)^(FadeTo|FadeBy|AppGain|AppMute)$") && (Type(val) = "Array")
            return this._vm.SetString(fullPath, "(" val[1] ", " val[2] ")")
		
		static toggles := "i)^(Mute|Solo|Mono|MC|A[1-5]|B[1-3]|On|PostReverb|PostDelay|PostFx1|PostFx2|Sel|Monitor)$"
        if (name ~= toggles) {
            current := this._vm.GetFloat(fullPath)
            ; If the user is trying to "Add" or "Subtract" (Math on the object)
            ; we ensure the result is always 0 or 1
            if (val > 1) || (val < 0)
                val := (current = 1 ? 0 : 1)
        }
        if IsNumber(val)
            return this._vm.SetFloat(fullPath, val)
        return this._vm.SetString(fullPath, val)
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
	static _instance := 0
		
	;Single Object reference only
	static Call(apiInstance := "") {
        if (this._instance == 0)
            this._instance := super.Call(apiInstance)
        return this._instance
    }
	
    __New(apiInstance) {
		if (apiInstance == "") {
            apiInstance := Voicemeeter()
        }
		if !(apiInstance is Voicemeeter) {
            throw TypeError("VoicemeeterProcess requires a valid Voicemeeter object. " 
                          . "Received: " . Type(apiInstance))
        }
        this.api := apiInstance
        this.pid := 0
        this.exe := ""
        this.hwnd := 0
		this._DetectExeFromDLL()
    }

    _DetectExeFromDLL() {
        if !this.api.EnsureConnected()
            return false

        old := DetectHiddenWindows(true)
        classes := ["VBCABLE0Voicemeeter0MainWindow0", "VMR_MainForm", "VBC_MainForm", "VBC8_MainForm"]

        for className in classes {
            if (hWin := WinExist("ahk_class " className)) {
                this.hwnd := hWin
                this.pid := WinGetPID(hWin)
                this.exe := WinGetProcessName(hWin)
                DetectHiddenWindows(old)
                return true
            }
        }

        static names := ["Voicemeeter", "VoicemeeterPro", "Voicemeeter8"]
        base := names[this.api.type + 1] 

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
        if (this.api.SetFloat("vban.Enable", newState)==0)
			return Float(newState)
		return this.api.GetFloat("vban.Enable")
    }

    RestartEngine(*) => this.api.SetFloat("Command.Restart", 1)

    RestartVoicemeeter(*) {
        this.Shutdown()
        this.hwnd := 0
        this.pid := 0
        return this.ShowOrHide()
    }
	
	Shutdown(*) => this.api.SetFloat("Command.Shutdown", 1)
	
    ShowOrHide(*) {
        if !this.exe
            this._DetectExeFromDLL()
        if !this.exe
            return false

        old := DetectHiddenWindows(true)

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

        if (hwnd := WinExist("ahk_exe " this.exe)) {
            this.hwnd := hwnd
            this.pid := WinGetPID(hwnd)
            DetectHiddenWindows(old)
            return this.ShowOrHide()
        }

        path := "C:\Program Files (x86)\VB\Voicemeeter\" this.exe
        try {
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