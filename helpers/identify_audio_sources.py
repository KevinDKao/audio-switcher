import sounddevice as sd

def list_audio_sources():
    print("Querying audio devices...")
    devices = sd.query_devices()
    
    print("\nAll Devices:")
    print(devices)

    print("\nOutput Devices (HostAPI 0 - usually MME on Windows, or looking for > 0 output channels):")
    # Filter for output devices (max_output_channels > 0)
    output_devices = [d for d in devices if d['max_output_channels'] > 0]
    
    for i, device in enumerate(output_devices):
        print(f"Index: {device['index']}, Name: {device['name']}, Channels: {device['max_output_channels']}, HostAPI: {device['hostapi']}")

if __name__ == "__main__":
    try:
        list_audio_sources()
    except Exception as e:
        print(f"Error: {e}")
