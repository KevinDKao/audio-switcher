import sys
import ctypes
from ctypes import wintypes
import comtypes
from comtypes import GUID

def main():
    try:
        from pycaw.pycaw import AudioUtilities
    except ImportError:
        print("Please install pycaw: pip install pycaw")
        return

    # Ensure COM is initialized for this thread
    ole32 = ctypes.windll.ole32
    ole32.CoInitialize(None)

    target_shure = "Headphones (2- Shure MV7+)"
    target_realtek = "Speakers (Realtek(R) Audio)"
    
    # 1. Device Discovery
    devices = AudioUtilities.GetAllDevices()
    
    device_map = {}
    for d in devices:
        if d.FriendlyName in [target_shure, target_realtek]:
            device_map[d.FriendlyName] = d.id
    
    if len(device_map) != 2:
        print(f"Error: Did not find both devices. Found: {list(device_map.keys())}")
        return

    id_shure = device_map[target_shure]
    id_realtek = device_map[target_realtek]

    # 2. Check current default
    default_device = AudioUtilities.GetSpeakers()
    default_id = default_device.id
    print(f"Current default ID: {default_id}")

    # 3. Switching Logic
    new_id = None
    new_name = ""
    
    if default_id == id_shure:
        print("Current is Shure. Switching to Realtek.")
        new_id = id_realtek
        new_name = target_realtek
    else:
        print("Current is not Shure. Switching to Shure.")
        new_id = id_shure
        new_name = target_shure
        
    s_new_id = new_id

    # 4. Set Default using manual VTable call via ctypes
    
    # Constants
    CLSID_PolicyConfig = GUID('{870af99c-171d-4f9e-af0d-e63df40c2bc9}')
    IID_IPolicyConfig = GUID('{f8679f50-850a-41cf-9c72-430f290290c8}')
    CLSCTX_ALL = 23
    
    CoCreateInstance = ole32.CoCreateInstance
    CoCreateInstance.argtypes = [
        ctypes.POINTER(GUID),
        ctypes.POINTER(ctypes.c_void_p), # pUnkOuter
        wintypes.DWORD, # dwClsContext
        ctypes.POINTER(GUID),
        ctypes.POINTER(ctypes.c_void_p) # ppv - changed to simpler definition
    ]
    CoCreateInstance.restype = ctypes.HRESULT
    
    p_policy_config = ctypes.c_void_p()
    # Pass address of p_policy_config directly
    hr = CoCreateInstance(
        ctypes.byref(CLSID_PolicyConfig),
        None,
        CLSCTX_ALL,
        ctypes.byref(IID_IPolicyConfig),
        ctypes.byref(p_policy_config)
    )
    
    if hr != 0:
        print(f"CoCreateInstance failed with 0x{hr:x}")
        ole32.CoUninitialize()
        return
        
    # Get VTable
    interface_ptr = p_policy_config.value
    if not interface_ptr:
        print("Interface pointer is null")
        ole32.CoUninitialize()
        return

    # Check pointer size for safety
    ptr_size = ctypes.sizeof(ctypes.c_void_p)
    
    # Read vtable_ptr
    vtable_ptr = ctypes.cast(interface_ptr, ctypes.POINTER(ctypes.c_void_p)).contents.value
    
    # Calculate offset for SetDefaultEndpoint (Index 13)
    # 0,1,2 IUnknown
    # 3..12 Others
    # 13 SetDefaultEndpoint
    func_offset = 13 * ptr_size
    func_address = vtable_ptr + func_offset
    
    # Read function pointer
    func_ptr = ctypes.cast(func_address, ctypes.POINTER(ctypes.c_void_p)).contents.value
    
    # Define function type
    SetDefaultEndpointType = ctypes.WINFUNCTYPE(
        ctypes.HRESULT,
        ctypes.c_void_p, # This
        ctypes.c_wchar_p, # DeviceID
        ctypes.c_int # Role
    )
    
    SetDefaultEndpoint = SetDefaultEndpointType(func_ptr)
    
    print(f"Setting default device to: {new_name}")
    
    success = True
    for role in [0, 1, 2]:
        try:
            hr = SetDefaultEndpoint(interface_ptr, s_new_id, role)
            if hr != 0:
                print(f"SetDefaultEndpoint failed for role {role}: 0x{hr:x}")
                success = False
        except Exception as e:
            print(f"Exception calling SetDefaultEndpoint: {e}")
            success = False

    # Cleanup
    # Release is index 2
    release_offset = 2 * ptr_size
    release_address = vtable_ptr + release_offset
    release_func_ptr = ctypes.cast(release_address, ctypes.POINTER(ctypes.c_void_p)).contents.value
    ReleaseType = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)
    Release = ReleaseType(release_func_ptr)
    Release(interface_ptr)
    
    ole32.CoUninitialize()
    
    if success:
        print(f"Switched to {new_name}")

if __name__ == "__main__":
    main()
