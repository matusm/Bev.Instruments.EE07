using System;
using System.IO.Ports;
using System.Threading;
using System.Collections.Generic;
using System.Text;

namespace Bev.Instruments.EE07
{
    public class EE07
    {
        private static SerialPort comPort;
        private const string genericString = "???";     // returned if something failed
        private const int numberTries = 20;             // number of tries before call gives up
        private const int delayTimeForRespond = 900;    // rather long delay nececssary
        // https://docs.microsoft.com/en-us/dotnet/api/system.io.ports.serialport.close?view=dotnet-plat-ext-5.0
        private const int waitOnClose = 50;             // No actual value is given, experimental

        private string cachedInstrumentType;
        private string cachedInstrumentSerialNumber;
        private string cachedInstrumentFirmwareVersion;


        public EE07(string portName)
        {
            DevicePort = portName.Trim();
            comPort = new SerialPort(DevicePort, 9600);
            comPort.RtsEnable = true;   // this is essential
            comPort.DtrEnable = true;	// this is essential
            ClearCache();
        }


        public string DevicePort { get; }
        public string InstrumentManufacturer => "E+E Elektronik";
        public string InstrumentType => GetInstrumentType();
        public string InstrumentSerialNumber => GetInstrumentSerialNumber();
        public string InstrumentFirmwareVersion => GetInstrumentVersion();
        public string InstrumentID => $"{InstrumentType} {InstrumentFirmwareVersion} SN:{InstrumentSerialNumber} @ {DevicePort}";
        public int SensorType { get; private set; }
        public double Temperature { get; private set; }
        public double Humidity { get; private set; }


        public void UpdateValues()
        {
            for (int i = 0; i < numberTries; i++)
            {
                _UpdateValues();
                if (!double.IsNaN(Temperature)) return;
            }
        }

        public void ClearCache()
        {
            cachedInstrumentType = genericString;
            cachedInstrumentSerialNumber = genericString;
            cachedInstrumentFirmwareVersion = genericString;
            SensorType = -1;
            ClearCachedValues();
        }

        private void ClearCachedValues()
        {
            Temperature = double.NaN;
            Humidity = double.NaN;
        }

        private string GetInstrumentType()
        {
            if (cachedInstrumentType == genericString)
                cachedInstrumentType = RepeatMethod(_GetInstrumentType);
            return cachedInstrumentType;
        }

        private string GetInstrumentVersion()
        {
            if (cachedInstrumentFirmwareVersion == genericString)
                cachedInstrumentFirmwareVersion = RepeatMethod(_GetInstrumentVersion);
            return cachedInstrumentFirmwareVersion;
        }

        private string GetInstrumentSerialNumber()
        {
            if (cachedInstrumentSerialNumber == genericString)
                cachedInstrumentSerialNumber = RepeatMethod(_GetInstrumentSerialNumber);
            return cachedInstrumentSerialNumber;
        }

        private string RepeatMethod(Func<string> getString)
        {
            for (int i = 0; i < numberTries; i++)
            {
                string str = getString();
                if (str != genericString)
                {
                    //if (i > 0) Console.WriteLine($"***** {i + 1} tries!");
                    return str;
                }
            }
            //Console.WriteLine($"***** {numberTries}, unsuccessfull!");
            return genericString;
        }

        private void _UpdateValues()
        {
            ClearCachedValues();
            // see E2Interface-RS232_e1.doc
            var reply = Query(0x58, new byte[] { 0x00, 0x30, 0x1E });
            if (reply.Length != 5)
            {
                return; // we need exactly 5 bytes
            }
            if (reply[4] != 0x00)
            {
                return; // if status gives an error, return
            }
            Humidity = (reply[0] + (reply[1]) * 256) / 100.0;
            Temperature = (reply[2] + reply[3] * 256) / 100.0 - 273.15;
        }

        private string _GetInstrumentType()
        {
            // undocumented!
            byte groupL, groupH, subGroup;
            byte[] reply;
            // Get group designation. 0x07 = EE07 etc.
            reply = Query(0x51, new byte[] { 0x11 });
            if (reply.Length != 1)
            {
                return genericString;
            }
            groupL = reply[0];
            // Get subgroup designation. 
            reply = Query(0x51, new byte[] { 0x21 });
            if (reply.Length != 1)
            {
                return genericString;
            }
            subGroup = reply[0];
            // Get group H-byte
            reply = Query(0x51, new byte[] { 0x41 });
            if (reply.Length != 1)
            {
                return genericString;
            }
            groupH = reply[0];
            // sensor type - what for?
            SensorType = groupH * 256 + groupL; //TODO make accessible?
            return $"EE{groupL:00}-{subGroup}";
        }

        private string _GetInstrumentVersion()
        {
            // undocumented!
            var reply = Query(0x55, new byte[] { 0x01, 0x80, 0x04 });
            if (reply.Length != 4)
                return genericString;
            var str = Encoding.UTF8.GetString(reply);
            str = str.Insert(2, ".");
            str = str.TrimStart('0');
            return str;
        }

        private string _GetInstrumentSerialNumber()
        {
            // undocumented!
            var reply = Query(0x55, new byte[] { 0x01, 0x84, 0x10 });
            if (reply.Length == 0)
                return genericString;
            // this is probably useless 
            // byte[] tempBuffer = Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(reply)); // this looks peculiar
            for (int i = 0; i < reply.Length; i++)
            {
                if (reply[i] == 0) reply[i] = 0x20; // substitute 0 by space
            }
            return Encoding.UTF8.GetString(reply).Trim();
        }

        private byte[] ComposeCommand(byte BField, byte[] DField)
        {
            List<byte> bufferList = new List<byte>();
            bufferList.Add(BField); // [B]
            if (DField == null || DField.Length == 0)
                bufferList.Add((byte)0);
            else
            {
                bufferList.Add((byte)DField.Length); // [L]
                foreach (byte b in DField)
                    bufferList.Add(b); // [D]
            }
            byte bsum = 0;
            foreach (byte b in bufferList)
                bsum += b;
            bufferList.Add(bsum); // [C]
            return bufferList.ToArray();
        }

        private byte[] Query(byte instruction, byte[] DField)
        {
            OpenPort();
            SendEE07(ComposeCommand(instruction, DField));
            Thread.Sleep(delayTimeForRespond);
            var buffer = ReadEE07();
            ClosePort();
            return AnalyzeRespond(buffer);
        }

        // This method takes the return byte array, checks if [L] is consistent,
        // if [S] is ACK and if the [CRC] is ok.
        // If so [C], [L], [S], [Sd] and [CRC] is stripped and the remaining array returned.
        private byte[] AnalyzeRespond(byte[] buffer)
        {
            var syntaxError = Array.Empty<byte>();
            if ((buffer.Length) < 5 || buffer == null)
            {
                // response too short
                return syntaxError;
            }
            // check CRC [C]
            byte bsum = 0;
            for (int i = 0; i < buffer.Length - 1; i++)
                bsum += buffer[i];
            if (bsum != buffer[buffer.Length - 1])
            {
                // CRC failed
                return syntaxError;
            }
            // check ACK
            if (buffer[2] != 0x06)
            {
                //TODO this is useless!
                if (buffer[2] == 0x15)
                {
                    // NAK
                    return syntaxError;
                }
                else
                {
                    // neither ACK nor NAK
                    return syntaxError;
                }
            }
            // check count of data bytes
            if (buffer[1] + 3 != buffer.Length)
            {
                // given data length not consistent
                return syntaxError;
            }
            byte[] tempbuff = new byte[buffer.Length - 5];
            for (int i = 4; i < buffer.Length - 1; i++)
                tempbuff[i - 4] = buffer[i];
            return tempbuff;
        }

        private void OpenPort()
        {
            try
            {
                if (!comPort.IsOpen)
                    comPort.Open();
            }
            catch (Exception)
            { }
        }

        private void ClosePort()
        {
            try
            {
                if (comPort.IsOpen)
                {
                    comPort.Close();
                    Thread.Sleep(waitOnClose);
                }
            }
            catch (Exception)
            { }
        }

        private void SendEE07(byte[] command)
        {
            try
            {
                comPort.Write(command, 0, command.Length);
                return;
            }
            catch (Exception)
            {
                //Console.WriteLine("***** SendEE07 failed: ", e);
                return;
            }
        }

        private byte[] ReadEE07()
        {
            byte[] ErrBuffer = { 0xFF };
            try
            {
                byte[] buffer = new byte[comPort.BytesToRead];
                comPort.Read(buffer, 0, buffer.Length);
                return buffer;
            }
            catch (Exception)
            {
                // Console.WriteLine("***** ReadEE07 failed: ", e);
                return ErrBuffer;
            }
        }

        // function for debbuging purposes
        private string BytesToString(byte[] bytes)
        {
            string str = "";
            foreach (byte b in bytes)
                str += $" {b,2:X2}";
            return str;
        }

    }
}
