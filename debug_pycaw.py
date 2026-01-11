from pycaw.pycaw import AudioUtilities
print("Pycaw imported successfully")
try:
    devices = AudioUtilities.GetAllDevices()
    print(f"Found {len(devices)} devices")
    for dev in devices:
        print(dev.FriendlyName)
except Exception as e:
    print(f"Error enumerating: {e}")
