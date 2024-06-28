using OneC.ExternalComponents;
using System;
using System.IO;
using System.Runtime.InteropServices;
using NAPS2.Wia;
using System.Linq;


namespace OneC.ExternalComponents {
	[Guid("7166906E-EAC8-43D1-9B32-4FFE8FDAEEC3")]
	[ProgId("AddIn.Com1CScan")]
	public class Com1CScan: ExtComponentBase {

		private const string Path = "C:\\Users\\Public\\hs.log";		
		private WiaDeviceManager _deviceManager;
		private uint _lastError = 0;

		public Com1CScan() {
			
			ComponentName = "_1CComScan";

			_deviceManager = new WiaDeviceManager();

		}

		~Com1CScan() {
			Disconnect();
		}		

		protected void WriteLog(string msg) {
			File.AppendAllText(Path, msg + Environment.NewLine);
		}

		[Export1c]
		public string GetLastError() {
			
			string result = "";

			switch(_lastError) {

				case 0x80210003:
					result = "There are no documents in the document feeder.";
					break;

				case 0x80210015:
					result = "No scanner device was found. Make sure the device is online, connected to the PC, and has the correct driver installed on the PC.";
					break;

				case 0x80210005:
					result = "The device is offline. Make sure the device is powered on and connected to the PC.";
					break;

				case 0x80210002:
					result = "Paper is jammed in the scanner's document feeder.";
					break;

				case 0x80210006:
					result = "The device is busy. Close any apps that are using this device or wait for it to finish and then try again.";				
					break;

				case 0x80210016:
					result = "One or more of the device’s cover is open.";
					break;

				case 0x8021000A:
					result = "ommunication with the WIA device failed. Make sure that the device is powered on and connected to the PC. If the problem persists, disconnect and reconnect the device.";
					break;

				case 0x8021000D:
					result = "The device is locked. Close any apps that are using this device or wait for it to finish and then try again.";
					break;

				case 0x8021000C:
					result = "There is an incorrect setting on the WIA device.";
					break;

				case 0x80210017:
					result = "The scanner's lamp is off.";
					break;

				case 0x80210007:
					result = "The device is warming up.";
					break;

				default:
					result = "Всё хорошо";
					break;
			}
			
			return result;
		}

		[Export1c]
		public int GetScanCount() {
			
			var infos = _deviceManager.GetDeviceInfos();			
			return infos.Count();

		}

		[Export1c]
		public string GetScanList() {

			try{
				
				WiaDevice device = _deviceManager.PromptForDevice();
				return device.Name() + "^;^" + device.Id();

			}catch(WiaException ex){
				_lastError = ex.ErrorCode;
			}			

			return "";

		}

		[Export1c]
		public bool StartScan(string id) {

			var _dev = _deviceManager.FindDevice(id);
			//_dev.FindSubItem
			return true;
		}

		[Export1c]
		public void Disconnect() {

		}

	}
}