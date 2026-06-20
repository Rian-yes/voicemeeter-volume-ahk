/*
    VOICEMEETER REMOTE API WRAPPER (AutoHotkey v2)
    @version : 2.1 (Merged Edition from VMR.ahk)
    @author : Rian
    -----------------------------------------------------------------------
    USAGE EXAMPLES:
    vm  := Voicemeeter()            ; Initialize (Singleton)
    vm1 := Voicemeeter()            ; Same object as vm object

    [ STRIPS & BUSES ]
    vm.strip[0].Mute := 1           ; Set Mute
    vm.strip[0].Solo += 1           ; Toggle Solo (Smart Toggle)
    vm.strip[1].Gain := -10.5       ; Set Gain (dB)
    vm.bus[0].Label := "Headset"    ; Set Label (String)
    
    [ SPECIAL COMMANDS ]
    vm.ShowOrHide()                 ; Toggle GUI visibility
    vm.RestartEngine()              ; Restart Audio Engine
    vm.command.Button[0].State := 1 ; Set Macro Button State
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
    
    static WindowClass => "ahk_class VBCABLE0Voicemeeter0MainWindow0"

    _type       := 0
    _lastType   := 0
    logged_in   := false
    inputs      := 0
    outputs     := 0
    _vmr        := 0
    
    static Call(Params*) {
        if !this._instance
            this._instance := super.Call(Params*)
        return this._instance
    }
    
    type {
        get {
            static input_map  := [0, 3, 5, 8]
            static output_map := [0, 2, 5, 8]
            rawType := this.GetVoicemeeterType()

            if (rawType > 0)
                this._lastType := rawType
            this._type := rawType

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
    
    __New(waitTimeoutMs := 15000) {
        this.connected := false
        this._writeCache := Map()
        this._vmr := Voicemeeter.RemoteInterface(Voicemeeter.DLL_PATH)
        this._Login()
        
        if (waitTimeoutMs > 0) {
            if (!this.WaitForServer(waitTimeoutMs)) {
                throw Voicemeeter.RemoteError("ERR_NO_SERVER", "__New", -2)
            }
        }
        
        _ := this.type
        
        ; Structural Nodes
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
        if this._vmr {
            this._Logout()
            this._vmr := 0
        }
        if this.HasOwnProp("_proc")
            this._proc.api := 0
        Voicemeeter._instance := 0
    }

    static _GetVoicemeeterDir() {
        static cachedPath := ""
        if (cachedPath != "")
            return cachedPath

        regView := A_RegView
        SetRegView 32
        uninstallString := RegRead("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Uninstall\VB:Voicemeeter {17359A74-1236-5467}", "UninstallString", "")
        SetRegView regView

        if (uninstallString != "") {
            SplitPath(uninstallString, , &dir)
            if DirExist(dir)
                return cachedPath := dir
        }
        return cachedPath := (EnvGet("ProgramFiles(x86)") || A_ProgramFiles) "\VB\Voicemeeter"
    }

    _Login() {
        res := DllCall(this._vmr.Login, "Int")
        if (res == 0 || res == 1) {
            return this.logged_in := true
        } else if (res == -2) {
            this._Logout()
            return this._Login()
        }
        throw Voicemeeter.RemoteError("ERR_UNEXPECTED", "_Login", res)
    }

    _Logout() {
        res := DllCall(this._vmr.Logout, "Int")
        this.logged_in := false
        if (res != 0)
            throw Voicemeeter.RemoteError("ERR_UNKNOWN", "_Logout", res)
        return res
    }

    GetVoicemeeterType() {
        val := Buffer(4)
        res := DllCall(this._vmr.GetVoicemeeterType, "Ptr", val, "Int")
        if (res == 0) 
			return NumGet(val, "Int")
        if (res == -2) 
			return 0
        throw Voicemeeter.RemoteError("ERR_UNEXPECTED", "GetVoicemeeterType", res)
    }

    GetVoicemeeterVersion() {
        val := Buffer(4)
        res := DllCall(this._vmr.GetVoicemeeterVersion, "Ptr", val, "Int")
        if (res == 0) 
			return NumGet(val, "Int")
        if (res == -2) 
			return 0
        throw Voicemeeter.RemoteError("ERR_UNEXPECTED", "GetVoicemeeterVersion", res)
    }

    EnsureConnected() {
        try {
            if (this.type > 0) {
                this.connected := true
                if (!this._proc.exe)
                    this._proc._DetectExeFromDLL()
                return true
            }
        }
        this.connected := false
        return false
    }

    WaitForServer(maxMs := 15000, sleepInterval := 100) {
        start := A_TickCount
        val := Buffer(4)
        Loop {
            res := DllCall(this._vmr.GetVoicemeeterType, "Ptr", val, "Int")
            if (res == 0) {
                if (DllCall(this._vmr.IsParametersDirty, "Int") >= 0) {
                    return true
                }
            }
            if (A_TickCount - start >= maxMs)
                return false
            Sleep sleepInterval
        }
    }

    IsParametersDirty() => DllCall(this._vmr.IsParametersDirty, "Int")

    WaitForNotDirty(maxMs := 500) {
        if (this.IsParametersDirty() == 0)
            return true
        start := A_TickCount
        Loop {
            if (this.IsParametersDirty() == 0)
                return true
            if (A_TickCount - start >= maxMs)
                return false
            Sleep 10
        }
    }

    GetFloat(p) => this.GetParameterFloat(p)
    SetFloat(p, v) => this.SetParameterFloat(p, v)
    GetString(p) => this.GetParameterString(p)
    SetString(p, v) => this.SetParameterString(p, v)

    GetParameterFloat(ParamName) {
        ; Check write-through cache first
        lowerName := StrLower(ParamName)
        if (this._writeCache.Has(lowerName)) {
            cache := this._writeCache[lowerName]
            if (A_TickCount - cache.time < 250) {
                this.WaitForNotDirty()
                val := Buffer(4)
                res := DllCall(this._vmr.GetParameterFloat, "AStr", ParamName, "Ptr", val, "Int")
                if (res == 0) {
                    dllVal := NumGet(val, "Float")
                    if (Abs(dllVal - cache.val) > 0.01) {
                        return cache.val
                    }
                }
            } else {
                this._writeCache.Delete(lowerName)
            }
        }

        ; Wait for parameter synchronization to complete
        this.WaitForNotDirty() 
        
        val := Buffer(4)
        Loop 10 {
            res := DllCall(this._vmr.GetParameterFloat, "AStr", ParamName, "Ptr", val, "Int")
            if (res == 0) 
                return NumGet(val, "Float")
            if (res != -1)
                break
            Sleep 50
        }
        
        errMap := Map(-2,"ERR_NO_SERVER", -3,"ERR_UNKNOWN_PARAMETER", -5,"ERR_STRUCTURE_MISMATCH")
        throw Voicemeeter.RemoteError(errMap.Get(res, "ERR_UNEXPECTED"), "GetParameterFloat", res)
    }

    SetParameterFloat(ParamName, Value) {
        Loop 10 {
            res := DllCall(this._vmr.SetParameterFloat, "AStr", ParamName, "Float", Float(Value), "Int")
            if (res == 0) {
                this._writeCache[StrLower(ParamName)] := {val: Float(Value), time: A_TickCount}
                return res
            }
            if (res != -1)
                break
            Sleep 50
        }
        errMap := Map(-2,"ERR_NO_SERVER", -3,"ERR_UNKNOWN_PARAMETER")
        throw Voicemeeter.RemoteError(errMap.Get(res, "ERR_UNEXPECTED"), "SetParameterFloat", res)
    }
    
    SetParameterString(ParamName, Value) {
        Loop 10 {
            res := DllCall(this._vmr.SetParameterString, "AStr", ParamName, "WStr", String(Value), "Int")
            if (res == 0) {
                this._writeCache[StrLower(ParamName)] := {val: String(Value), time: A_TickCount}
                return res
            }
            if (res != -1)
                break
            Sleep 50
        }
        errMap := Map(-2,"ERR_NO_SERVER", -3,"ERR_UNKNOWN_PARAMETER")
        throw Voicemeeter.RemoteError(errMap.Get(res, "ERR_UNEXPECTED"), "SetParameterString", res)
    }

    GetParameterString(ParamName) {
        lowerName := StrLower(ParamName)
        if (this._writeCache.Has(lowerName)) {
            cache := this._writeCache[lowerName]
            if (A_TickCount - cache.time < 250) {
                this.WaitForNotDirty()
                buf := Buffer(1024, 0)
                res := DllCall(this._vmr.GetParameterString, "AStr", ParamName, "Ptr", buf, "Int")
                if (res == 0) {
                    dllVal := StrGet(buf, "UTF-16")
                    if (dllVal != cache.val) {
                        return cache.val
                    }
                }
            } else {
                this._writeCache.Delete(lowerName)
            }
        }

        this.WaitForNotDirty()
        buf := Buffer(1024, 0)
        Loop 10 {
            res := DllCall(this._vmr.GetParameterString, "AStr", ParamName, "Ptr", buf, "Int")
            if (res == 0) 
                return StrGet(buf, "UTF-16")
            if (res != -1)
                break
            Sleep 50
        }
        errMap := Map(-2,"ERR_NO_SERVER", -3,"ERR_UNKNOWN_PARAMETER", -5,"ERR_STRUCTURE_MISMATCH")
        throw Voicemeeter.RemoteError(errMap.Get(res, "ERR_UNEXPECTED"), "GetParameterString", res)
    }

    SetParameters(Params) {
        Loop 10 {
            res := DllCall(this._vmr.SetParameters, "WStr", String(Params), "Int")
            if (res == 0) 
                return res
            if (res != -1)
                break
            Sleep 50
        }
        if (res > 0) 
            throw Voicemeeter.RemoteError("ERR_SCRIPT_ERROR", "SetParameters", res, Params)
        errMap := Map(-2,"ERR_NO_SERVER")
        throw Voicemeeter.RemoteError(errMap.Get(res, "ERR_UNEXPECTED"), "SetParameters", res)
    }

    GetLevel(type, channel) {
        val := Buffer(4)
        res := DllCall(this._vmr.GetLevel, "Int", type, "Int", channel, "Ptr", val, "Int")
        if (res == 0) 
			return NumGet(val, "Float")
        errMap := Map(-2,"ERR_NO_SERVER", -3,"ERR_NO_LEVEL_AVAILABLE", -4,"ERR_OUT_OF_RANGE")
        throw Voicemeeter.RemoteError(errMap.Get(res, "ERR_UNEXPECTED"), "GetLevel", res)
    }

    GetMidiMessage(&MidiBuffer, maxSize := 1024) {
        MidiBuffer := Buffer(maxSize, 0)
        res := DllCall(this._vmr.GetMidiMessage, "Ptr", MidiBuffer, "Int", maxSize, "Int")
        if (res >= 0) 
			return res
        errMap := Map(-2,"ERR_NO_SERVER", -5,"ERR_NO_MIDI_DATA")
        throw Voicemeeter.RemoteError(errMap.Get(res, "ERR_UNEXPECTED"), "GetMidiMessage", res)
    }

    SendMidiMessage(data, size?) {
        if (Type(data) = "Array" || data is Array) {
            sz := IsSet(size) ? size : data.Length
            buf := Buffer(sz, 0)
            for idx, byte in data {
                if (idx > sz)
                    break
                NumPut("UChar", byte, buf, idx - 1)
            }
            ptr := buf.Ptr
        } else if (data is Integer) {
            sz := IsSet(size) ? size : 4
            ptr := data
        } else {
            sz := IsSet(size) ? size : data.Size
            ptr := data.Ptr
        }
        res := DllCall(this._vmr.SendMidiMessage, "Ptr", ptr, "Int", sz, "Int")
        if (res == 0) 
			return res
        errMap := Map(-2,"ERR_NO_SERVER", -6,"ERR_CANNOT_SEND_MIDI_DATA")
        throw Voicemeeter.RemoteError(errMap.Get(res, "ERR_UNEXPECTED"), "SendMidiMessage", res)
    }

    MacroButton_IsDirty() {
        res := DllCall(this._vmr.MacroButton_IsDirty, "Int")
        if (res >= 0) 
			return res
        errMap := Map(-2,"ERR_NO_SERVER")
        throw Voicemeeter.RemoteError(errMap.Get(res, "ERR_UNEXPECTED"), "MacroButton_IsDirty", res)
    }

    MacroButton_GetStatus(logicalButton, bitmode := 0) {
        val := Buffer(4)
        res := DllCall(this._vmr.MacroButton_GetStatus, "Int", logicalButton, "Ptr", val, "Int", bitmode, "Int")
        if (res == 0) 
			return NumGet(val, "Float")
        errMap := Map(-2,"ERR_NO_SERVER", -3,"ERR_OUT_OF_RANGE")
        throw Voicemeeter.RemoteError(errMap.Get(res, "ERR_UNEXPECTED"), "MacroButton_GetStatus", res)
    }

    MacroButton_SetStatus(logicalButton, value, bitmode := 0) {
        res := DllCall(this._vmr.MacroButton_SetStatus, "Int", logicalButton, "Float", Float(value), "Int", bitmode, "Int")
        if (res == 0) 
			return res
        errMap := Map(-2,"ERR_NO_SERVER", -3,"ERR_OUT_OF_RANGE")
        throw Voicemeeter.RemoteError(errMap.Get(res, "ERR_UNEXPECTED"), "MacroButton_SetStatus", res)
    }

    AudioCallbackRegister(mode, callbackAddress, userPtr := 0, clientName := "AHK_Voicemeeter") {
        res := DllCall(this._vmr.AudioCallbackRegister, "Int", mode, "Ptr", callbackAddress, "Ptr", userPtr, "AStr", clientName, "Int")
        if (res == 0) 
			return res
        errMap := Map(-2,"ERR_NO_SERVER")
        throw Voicemeeter.RemoteError(errMap.Get(res, "ERR_UNEXPECTED"), "AudioCallbackRegister", res)
    }

    AudioCallbackStart() {
        res := DllCall(this._vmr.AudioCallbackStart, "Int")
        if (res == 0) 
			return res
        errMap := Map(-2,"ERR_NO_SERVER")
        throw Voicemeeter.RemoteError(errMap.Get(res, "ERR_UNEXPECTED"), "AudioCallbackStart", res)
    }

    AudioCallbackStop() {
        res := DllCall(this._vmr.AudioCallbackStop, "Int")
        if (res == 0) 
			return res
        errMap := Map(-2,"ERR_NO_SERVER")
        throw Voicemeeter.RemoteError(errMap.Get(res, "ERR_UNEXPECTED"), "AudioCallbackStop", res)
    }

    AudioCallbackUnregister() {
        res := DllCall(this._vmr.AudioCallbackUnregister, "Int")
        if (res == 0) 
			return res
        errMap := Map(-2,"ERR_NO_SERVER")
        throw Voicemeeter.RemoteError(errMap.Get(res, "ERR_UNEXPECTED"), "AudioCallbackUnregister", res)
    }

    GetOutputDeviceCount() => DllCall(this._vmr.Output_GetDeviceNumber, "Int")

    GetOutputDeviceDescriptor(Index) {
        deviceType := Buffer(4)
        deviceName := Buffer(512)
        hardwareId := Buffer(512)
        res := DllCall(this._vmr.Output_GetDeviceDesc, "Int", Index, "Ptr", deviceType, "Ptr", deviceName, "Ptr", hardwareId, "Int")
        if (res != 0)
            throw Voicemeeter.RemoteError("ERR_UNKNOWN", "GetOutputDeviceDescriptor", res)
        return Voicemeeter.DeviceDescriptor(Index, NumGet(deviceType, "Int"), StrGet(deviceName, "UTF-16"), StrGet(hardwareId, "UTF-16"))
    }

    GetOutputDeviceDescriptors() {
        descriptors := []
        count := this.GetOutputDeviceCount()
        Loop count {
            descriptors.Push(this.GetOutputDeviceDescriptor(A_Index - 1))
        }
        return descriptors
    }

    GetInputDeviceCount() => DllCall(this._vmr.Input_GetDeviceNumber, "Int")

    GetInputDeviceDescriptor(Index) {
        deviceType := Buffer(4)
        deviceName := Buffer(512)
        hardwareId := Buffer(512)
        res := DllCall(this._vmr.Input_GetDeviceDesc, "Int", Index, "Ptr", deviceType, "Ptr", deviceName, "Ptr", hardwareId, "Int")
        if (res != 0)
            throw Voicemeeter.RemoteError("ERR_UNKNOWN", "GetInputDeviceDescriptor", res)
        return Voicemeeter.DeviceDescriptor(Index, NumGet(deviceType, "Int"), StrGet(deviceName, "UTF-16"), StrGet(hardwareId, "UTF-16"))
    }

    GetInputDeviceDescriptors() {
        descriptors := []
        count := this.GetInputDeviceCount()
        Loop count {
            descriptors.Push(this.GetInputDeviceDescriptor(A_Index - 1))
        }
        return descriptors
    }

    BuildParamString(Value*) {
        str := ""
        for i in Value {
            str .= Value[i] . ";"
        }
        return SubStr(str, 1, -1)
    }

    ShowVoicemeeterWindow() {
        WinShow Voicemeeter.WindowClass
        WinActivate Voicemeeter.WindowClass
    }

    HideVoicemeeterWindow() {
        WinHide Voicemeeter.WindowClass
    }

    ToggleVoicemeeterWindow() {
        if WinActive(Voicemeeter.WindowClass) {
            this.HideVoicemeeterWindow()
        } else {
            this.ShowVoicemeeterWindow()
        }
    }

    class DeviceType {
        static MME  := 1
        static WDM  := 3
        static KS   := 4
        static ASIO := 5
    }

    class DeviceDescriptor {
        __New(Index, DeviceType, Name, HardwareId) {
            this.Index := Index
            this.Type := DeviceType
            this.Name := Name
            this.HardwareId := HardwareId
        }
    }

    ; Custom error engine wrapper derived from Antigravity's setup
    class RemoteError extends Error {
        Code := 0
        __New(ErrorType, What, Code?, Extra?) {
            prefix := IsSet(Code) ? "VBVMR (" . Code . "): " : "VBVMR: "
            errMsgs := Map(
                "ERR_NOT_INSTALLED", "Voicemeeter is not installed.",
                "ERR_UNKNOWN_VTYPE", "Unknown Voicemeeter type number.",
                "ERR_UNEXPECTED", "An unexpected error occurred.",
                "ERR_NO_SERVER", "Server not found.",
                "ERR_UNKNOWN_PARAMETER", "Unknown parameter.",
                "ERR_STRUCTURE_MISMATCH", "Structure mismatch.",
                "ERR_NO_LEVEL_AVAILABLE", "No level available.",
                "ERR_OUT_OF_RANGE", "Out of range.",
                "ERR_NO_MIDI_DATA", "No MIDI data.",
                "ERR_CANNOT_SEND_MIDI_DATA", "Cannot send MIDI data."
            )
            if (ErrorType == "ERR_SCRIPT_ERROR") {
                message := IsSet(Code) ? "Script contains an error on line " . Code . "." : "Script contains an error."
            } else {
                message := errMsgs.Get(ErrorType, "An unknown error occurred.")
            }
            super.__New(prefix . message, What, Extra?)
            if IsSet(Code)
                this.Code := Code
        }
    }

    class RemoteInterface {
        __New(DllPath) {
            this._hModule := DllCall("LoadLibrary", "Str", DllPath, "Ptr")
            if !this._hModule
                throw Error("Failed to load Voicemeeter DLL: " DllPath)

            GetProc(ProcName) {
                ptr := DllCall("GetProcAddress", "Ptr", this._hModule, "AStr", ProcName, "Ptr")
                if !ptr
                    throw Error("Failed to resolve DLL export: " ProcName)
                return ptr
            }

            this.Login                  := GetProc("VBVMR_Login")
            this.Logout                 := GetProc("VBVMR_Logout")
            this.RunVoicemeeter         := GetProc("VBVMR_RunVoicemeeter")
            this.GetVoicemeeterType     := GetProc("VBVMR_GetVoicemeeterType")
            this.GetVoicemeeterVersion  := GetProc("VBVMR_GetVoicemeeterVersion")
            this.IsParametersDirty      := GetProc("VBVMR_IsParametersDirty")
            this.GetParameterFloat      := GetProc("VBVMR_GetParameterFloat")
            this.GetParameterString     := GetProc("VBVMR_GetParameterStringW")
            this.GetLevel               := GetProc("VBVMR_GetLevel")
            this.GetMidiMessage         := GetProc("VBVMR_GetMidiMessage")
            this.SetParameterFloat      := GetProc("VBVMR_SetParameterFloat")
            this.SetParameterString     := GetProc("VBVMR_SetParameterStringW")
            this.SetParameters          := GetProc("VBVMR_SetParametersW")
            this.Output_GetDeviceNumber := GetProc("VBVMR_Output_GetDeviceNumber")
            this.Output_GetDeviceDesc   := GetProc("VBVMR_Output_GetDeviceDescW")
            this.Input_GetDeviceNumber  := GetProc("VBVMR_Input_GetDeviceNumber")
            this.Input_GetDeviceDesc    := GetProc("VBVMR_Input_GetDeviceDescW")
            this.MacroButton_IsDirty    := GetProc("VBVMR_MacroButton_IsDirty")
            this.MacroButton_GetStatus  := GetProc("VBVMR_MacroButton_GetStatus")
            this.MacroButton_SetStatus  := GetProc("VBVMR_MacroButton_SetStatus")
            this.SendMidiMessage        := GetProc("VBVMR_SendMidiMessage")
            
            ; Audio Callbacks
            this.AudioCallbackRegister   := GetProc("VBVMR_AudioCallbackRegister")
            this.AudioCallbackStart      := GetProc("VBVMR_AudioCallbackStart")
            this.AudioCallbackStop       := GetProc("VBVMR_AudioCallbackStop")
            this.AudioCallbackUnregister := GetProc("VBVMR_AudioCallbackUnregister")
        }
        __Delete() => DllCall("FreeLibrary", "Ptr", this._hModule)
    }
}

;==================================================================
; Optimized VMNode System
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

        ; OPTIMIZATION: High-performance structural container checks
        static containers := Map(
            "comp",1,"gate",1,"denoiser",1,"pitch",1,"gainlayer",1,"eq",1,"device",1,
            "app",1,"mode",1,"patch",1,"color",1,"fx",1,"outstream",1,"instream",1,
            "buffer",1,"button",1,"preset",1,"dialogshow",1,"armstrip",1,"armbus",1,
            "item",1,"delay",1,"sr",1,"slim",1,"wdm",1,"ks",1,"mme",1,"asio",1,
            "channel",1,"cell",1,"reverb",1
        )
        if (params.Length > 0 || containers.Has(StrLower(name)))
            return VMNode(this._vm, fullPath)
            
        static stringParams := Map(
            "label",1,"name",1,"ip",1,"sr",1,"channel",1,"bit",1,
            "fadeto",1,"fadeby",1,"appgain",1,"appmute",1,"goto",1,"load",1,"save",1,
            "loadbuseq",1,"savebuseq",1,"loadstripeq",1,"savestripeq",1
        )
        if (stringParams.Has(StrLower(name)))
            return this._vm.GetParameterString(fullPath)

        return this._vm.GetParameterFloat(fullPath)
    }

    __Set(name, params, val) {
        if (StrLower(name) == "gain" && IsNumber(val)) {
            val := (val > 12) ? 12 : (val < -60) ? -60 : val
        }
        part := name
        for p in params
            part .= "[" p "]"
        fullPath := this._prefix (this._prefix ? "." : "") part

        static arrays := Map("fadeto",1,"fadeby",1,"appgain",1,"appmute",1)
        if (arrays.Has(StrLower(name)) && (Type(val) = "Array" || val is Array))
            return this._vm.SetParameterString(fullPath, "(" val[1] ", " val[2] ")")

        static writeStrings := Map(
            "save",1,"load",1,"savebuseq",1,"loadbuseq",1,"savestripeq",1,"loadstripeq",1,
            "label",1,"name",1,"ip",1,"filenameattr",1,"goto",1,"wdm",1,"ks",1,"mme",1,"asio",1
        )
        if (writeStrings.Has(StrLower(name)) || (!IsNumber(val) && Type(val) = "String"))
            return this._vm.SetParameterString(fullPath, val)

        static toggles := Map(
            "mute",1, "solo",1, "mono",1, "mc",1, "on",1, "postreverb",1, "postdelay",1, 
            "postfx1",1, "postfx2",1, "sel",1, "monitor",1, "lock",1, "eject",1, "reset",1, 
            "show",1, "shutdown",1, "record",1, "play",1, "stop",1, "loop",1, "makeup",1, 
            "state",1, "stateonly",1, "trigger",1, "recall",1, "restart",1, "enable",1
        )
        if (toggles.Has(StrLower(name)) || name ~= "i)^([AB][1-5]|EQ\.On)$") {
            current := this._vm.GetParameterFloat(fullPath)
            if (val > 1 || val < 0)
                val := (current = 1 ? 0 : 1)
        }
        return this._vm.SetParameterFloat(fullPath, val)
    }

    __Cast(target) {
        if (target = "Number" || target = "Float" || target = "Integer")
            return this._vm.GetParameterFloat(this._prefix)
        return this._vm.GetParameterString(this._prefix)
    }
}

class VoicemeeterProcess {
    exe  := ""
    pid  := 0
    hwnd := 0
    
    __New(apiInstance) {
        this.api := apiInstance ? apiInstance : Voicemeeter()
        this._DetectExeFromDLL()
    }

    _DetectExeFromDLL() {
        if (this.api.type == 0) 
			return false
        this.api.connected := true

        old := DetectHiddenWindows(true)
        if (hWin := WinExist(Voicemeeter.WindowClass)) {
            this.hwnd := hWin
            this.pid := WinGetPID(hWin)
            this.exe := WinGetProcessName(hWin)
            DetectHiddenWindows(old)
            return true
        }

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
            return this.api.GetParameterFloat("vban.Enable")
        if (this.api.SetParameterFloat("vban.Enable", newState) == 0)
            return Float(newState)
        return this.api.GetParameterFloat("vban.Enable")
    }

    RestartEngine(*) => this.api.SetParameterFloat("Command.Restart", 1)

    RestartVoicemeeter(*) {
        this.Shutdown()
        if (this.exe) {
            try ProcessWaitClose(this.exe, 3)
        }
        this.hwnd := 0
        this.pid := 0
        return this.ShowOrHide()
    }
    
    Shutdown(*) => this.api.SetParameterFloat("Command.Shutdown", 1)
    
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
            DllCall(this.api._vmr.RunVoicemeeter, "Int", vType, "Int")
            
            if (this.hwnd := WinWait("ahk_exe " this.exe, , 5)) {
                this.pid := WinGetPID(this.hwnd)
                WinShow("ahk_id " this.hwnd)
                WinActivate("ahk_id " this.hwnd)
                DetectHiddenWindows(old)
                return true
            }
            DetectHiddenWindows(old)
            return false
        }

        ; Case 4: No DLL launch available — run the exe directly
        path := Voicemeeter.File_Dir "\" this.exe
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
