; #Include E:\ahk\VA.ahk
#Include VA.ahk

F1::  ; F1 hotkey - toggle mute state of active window
  WindowEXE := WinExist("A")
    ControlGetFocus, FocusedControl, ahk_id %WindowEXE%
    ControlGet, Hwnd, Hwnd,, %FocusedControl%, ahk_id %WindowEXE%
    WinGet, simplexe, processname, ahk_id %Hwnd%
  if !(Volume := GetVolumeObject(simplexe))
   MsgBox, There was a problem retrieving the application volume interface
return

;Required for app specific mute
GetVolumeObject(ProcessName) {
    static IID_IASM2 := "{77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}"
    , IID_IASC2 := "{bfb7ff88-7239-4fc9-8fa2-07c950be9c6d}"
    , IID_ISAV := "{87CE5498-68D6-44E5-9215-6DA47EF883D8}"

    ; Initialize an array to store ISAV objects
    ISAVArray := []

    ; Get all audio devices
    Loop, 10 ; Change the loop limit based on the number of audio devices you have
    {
        DAE := VA_GetDevice(A_Index)
        if (DAE)
        {
            ; Activate the session manager
            VA_IMMDevice_Activate(DAE, IID_IASM2, 0, 0, IASM2)

            ; Enumerate sessions for the current device
            VA_IAudioSessionManager2_GetSessionEnumerator(IASM2, IASE)
            VA_IAudioSessionEnumerator_GetCount(IASE, Count)

            ; Search for all instances of the specified process name for the current device
            Loop, % Count
            {
                VA_IAudioSessionEnumerator_GetSession(IASE, A_Index-1, IASC)
                IASC2 := ComObjQuery(IASC, IID_IASC2)

                ; If IAudioSessionControl2 is queried successfully
                if (IASC2)
                {
                    VA_IAudioSessionControl2_GetProcessID(IASC2, SPID)
                    ProcessNameFromPID := GetProcessNameFromPID(SPID)

                    ; If the process name matches the one we are looking for
                    if (ProcessNameFromPID == ProcessName)
                    {
                        ISAV := ComObjQuery(IASC2, IID_ISAV)
                        ISAVArray.Insert(ISAV)
                    }

                    ObjRelease(IASC2)
                }

                ObjRelease(IASC)
            }

            ObjRelease(IASE)
            ObjRelease(IASM2)
            ObjRelease(DAE)
        }
    }

    ; Mute all found instances
    Loop, % ISAVArray.Length()
    {
        VA_ISimpleAudioVolume_GetMute(ISAVArray[A_Index-1], Mute)
        VA_ISimpleAudioVolume_SetMute(ISAVArray[A_Index-1], !Mute)
        ObjRelease(ISAVArray[A_Index-1])
    }

    return ISAVArray  ; Return the array of ISAV objects
}

GetProcessNameFromPID(PID)
{
    hProcess := DllCall("OpenProcess", "UInt", 0x0400 | 0x0010, "Int", false, "UInt", PID)
    VarSetCapacity(ExeName, 260, 0)
    DllCall("Psapi.dll\GetModuleFileNameEx", "UInt", hProcess, "UInt", 0, "Str", ExeName, "UInt", 260)
    DllCall("CloseHandle", "UInt", hProcess)
    return SubStr(ExeName, InStr(ExeName, "\", false, -1) + 1)
}

;
; ISimpleAudioVolume : {87CE5498-68D6-44E5-9215-6DA47EF883D8}
;
VA_ISimpleAudioVolume_SetMasterVolume(this, ByRef fLevel, GuidEventContext="") {
    return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "float", fLevel, "ptr", VA_GUID(GuidEventContext))
}
VA_ISimpleAudioVolume_GetMasterVolume(this, ByRef fLevel) {
    return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "float*", fLevel)
}
VA_ISimpleAudioVolume_SetMute(this, ByRef Muted, GuidEventContext="") {
    return DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "int", Muted, "ptr", VA_GUID(GuidEventContext))
}
VA_ISimpleAudioVolume_GetMute(this, ByRef Muted) {
    return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "int*", Muted)
}
