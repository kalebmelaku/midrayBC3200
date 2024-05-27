namespace Mindray_BC3200
{
    using Dapper;
    using System;
    using System.ComponentModel;
    using System.Data;
    using System.Data.SqlClient;
    using System.IO;
    using System.IO.Ports;
    using System.Security.Cryptography;
    using System.ServiceProcess;
    using System.Text;

    public class Service1 : ServiceBase
    {
        private SqlConnection con = null;
        private SerialPort serialPort = null;
        private string fullMSG = "";
        private static string hash = "Nehemiahz@123";
        private static string LisPath = @"C:\Medaxs\LIS_Path.txt";
        private static string logFilePath = @"C:\Medaxs\Mindray_BC3200\Mindray_BC3200.txt";
        private static string logPath = @"C:\Medaxs\Mindray_BC3200";
        private static string COMSettingPath = @"C:\Medaxs\COMSetting\COMSetting.txt";
        private IContainer components = null;

        public Service1()
        {
            this.InitializeComponent();
        }

        public string Decryptor(string text)
        {
            string str;
            try
            {
                byte[] inputBuffer = Convert.FromBase64String(text);
                using (MD5CryptoServiceProvider provider = new MD5CryptoServiceProvider())
                {
                    byte[] buffer2 = provider.ComputeHash(Encoding.UTF8.GetBytes(hash));
                    TripleDESCryptoServiceProvider provider1 = new TripleDESCryptoServiceProvider();
                    provider1.Key = buffer2;
                    provider1.Mode = CipherMode.ECB;
                    provider1.Padding = PaddingMode.PKCS7;
                    using (TripleDESCryptoServiceProvider provider2 = provider1)
                    {
                        byte[] bytes = provider2.CreateDecryptor().TransformFinalBlock(inputBuffer, 0, inputBuffer.Length);
                        str = Encoding.UTF8.GetString(bytes);
                    }
                }
            }
            catch (Exception exception)
            {
                this.TakeLog("Exception on Decrypting the DB Path: " + exception.Message.ToString());
                str = "";
            }
            return str;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void ExtractData(string x)
        {
            try
            {
                if (!(x.StartsWith("\x0002") && x.EndsWith("\x001a")))
                {
                    this.TakeLog("Wrong Formattt!");
                }
                else
                {
                    int num = Convert.ToInt32(x.Substring(15, 7));
                    string str2 = "WBC";
                    string str3 = "LYM#";
                    string str4 = "MID#";
                    string str5 = "Gran#";
                    string str6 = "LYM%";
                    string str7 = "Mid%";
                    string str8 = "Gran%";
                    string str9 = "RBC";
                    string str10 = "HGB";
                    string str11 = "MCHC";
                    string str12 = "MCV";
                    string str13 = "MCH";
                    string str14 = "RDW-CV";
                    string str15 = "HCT";
                    string str16 = "PLT";
                    string str17 = "MPV";
                    string str18 = "PDW";
                    string str19 = "PCT";
                    string str20 = "RDW-SD";
                    string str21 = "";
                    string str22 = "";
                    string str23 = "";
                    string str24 = "";
                    string str25 = "";
                    string str26 = "";
                    string str27 = "";
                    string str28 = "";
                    string str29 = "";
                    string str30 = "";
                    string str31 = "";
                    string str32 = "";
                    string str33 = "";
                    string str34 = "";
                    string str35 = "";
                    string str36 = "";
                    string str37 = "";
                    string str38 = "";
                    string str39 = "";
                    str23 = Convert.ToInt32(x.Substring(0x24, 2)).ToString() + "." + x.Substring(0x26, 1);
                    str24 = Convert.ToInt32(x.Substring(0x29, 1)).ToString() + "." + x.Substring(0x2a, 1);
                    str25 = Convert.ToInt32(x.Substring(0x2d, 1)).ToString() + "." + x.Substring(0x2e, 1);
                    str26 = Convert.ToInt32(x.Substring(0x30, 2)).ToString() + "." + x.Substring(50, 1);
                    str27 = Convert.ToInt32(x.Substring(0x33, 2)).ToString() + "." + x.Substring(0x35, 1);
                    str28 = Convert.ToInt32(x.Substring(0x36, 2)).ToString() + "." + x.Substring(0x38, 1);
                    str29 = Convert.ToInt32(x.Substring(0x39, 2)).ToString() + "." + x.Substring(0x3b, 1);
                    str30 = Convert.ToInt32(x.Substring(60, 1)).ToString() + "." + x.Substring(0x3d, 2);
                    str31 = Convert.ToInt32(x.Substring(0x3f, 2)).ToString() + "." + x.Substring(0x41, 1);
                    str32 = Convert.ToInt32(x.Substring(0x43, 2)).ToString() + "." + x.Substring(0x45, 1);
                    str33 = Convert.ToInt32(x.Substring(70, 3)).ToString() + "." + x.Substring(0x49, 1);
                    str34 = Convert.ToInt32(x.Substring(0x4b, 2)).ToString() + "." + x.Substring(0x4d, 1);
                    str35 = Convert.ToInt32(x.Substring(0x4e, 2)).ToString() + "." + x.Substring(80, 1);
                    str36 = Convert.ToInt32(x.Substring(0x51, 2)).ToString() + "." + x.Substring(0x53, 1);
                    str21 = Convert.ToInt32(x.Substring(0x55, 3)).ToString() ?? "";
                    str37 = Convert.ToInt32(x.Substring(0x58, 2)).ToString() + "." + x.Substring(90, 1);
                    str38 = Convert.ToInt32(x.Substring(0x5b, 2)).ToString() + "." + x.Substring(0x5d, 1);
                    str39 = "0." + Convert.ToInt32(x.Substring(0x5e, 3)).ToString();
                    str22 = Convert.ToInt32(x.Substring(0x62, 2)).ToString() + "." + x.Substring(100, 1);
                    this.TakeLog(str2 + ": " + str23);
                    this.TakeLog(str3 + ": " + str24);
                    this.TakeLog(str4 + ": " + str25);
                    this.TakeLog(str5 + ": " + str26);
                    this.TakeLog(str6 + ": " + str27);
                    this.TakeLog(str7 + ": " + str28);
                    this.TakeLog(str8 + ": " + str29);
                    this.TakeLog(str9 + ": " + str30);
                    this.TakeLog(str10 + ": " + str31);
                    this.TakeLog(str11 + ": " + str32);
                    this.TakeLog(str12 + ": " + str33);
                    this.TakeLog(str13 + ": " + str34);
                    this.TakeLog(str14 + ": " + str35);
                    this.TakeLog(str15 + ": " + str36);
                    this.TakeLog(str16 + ": " + str21);
                    this.TakeLog(str17 + ": " + str37);
                    this.TakeLog(str18 + ": " + str38);
                    this.TakeLog(str19 + ": " + str39);
                    this.TakeLog(str20 + ": " + str22);
                    try
                    {
                        using (IDbConnection connection = DBConnection.GetConnection())
                        {
                            string[] textArray1 = new string[0x29];
                            textArray1[0] = "INSERT INTO  `cbc` (`patient_id` , `WBC`, `LYM#`, `MID#`, `GRA#`, `LYM%`, `MID%`, `GRA%`, `RBC`, `HGB`, `MCHC`,  `MCV`, `MCH`, `RDW-CV`,  `HCT` ,  `PLT`, `MPV`, `PDW`, `PCT`, `RDW-SD`) VALUES ('";
                            textArray1[1] = num.ToString();
                            textArray1[2] = "','";
                            textArray1[3] = str23;
                            textArray1[4] = "','";
                            textArray1[5] = str24;
                            textArray1[6] = "','";
                            textArray1[7] = str25;
                            textArray1[8] = "','";
                            textArray1[9] = str26;
                            textArray1[10] = "','";
                            textArray1[11] = str27;
                            textArray1[12] = "','";
                            textArray1[13] = str28;
                            textArray1[14] = "','";
                            textArray1[15] = str29;
                            textArray1[0x10] = "','";
                            textArray1[0x11] = str30;
                            textArray1[0x12] = "','";
                            textArray1[0x13] = str31;
                            textArray1[20] = "','";
                            textArray1[0x15] = str32;
                            textArray1[0x16] = "','";
                            textArray1[0x17] = str33;
                            textArray1[0x18] = "','";
                            textArray1[0x19] = str34;
                            textArray1[0x1a] = "','";
                            textArray1[0x1b] = str35;
                            textArray1[0x1c] = "','";
                            textArray1[0x1d] = str36;
                            textArray1[30] = "','";
                            textArray1[0x1f] = str21;
                            textArray1[0x20] = "','";
                            textArray1[0x21] = str37;
                            textArray1[0x22] = "','";
                            textArray1[0x23] = str38;
                            textArray1[0x24] = "','";
                            textArray1[0x25] = str39;
                            textArray1[0x26] = "','";
                            textArray1[0x27] = str22;
                            textArray1[40] = "')";
                            string sql = string.Concat(textArray1);
                            if (connection.State != ConnectionState.Open)
                            {
                                connection.Open();
                            }
                            int? commandTimeout = null;
                            CommandType? commandType = null;
                            connection.Execute(sql, new <>f__AnonymousType0(), null, commandTimeout, commandType);
                            this.TakeLog("Successfully Inserted");
                        }
                    }
                    catch (Exception exception)
                    {
                        this.TakeLog("Exception Occured: " + exception.Message);
                    }
                }
            }
            catch (Exception exception2)
            {
                this.TakeLog("Exception Ocurred in ExtractData() Method: " + exception2.Message.ToString());
            }
        }

        private void InitializeComponent()
        {
            this.components = new Container();
            base.ServiceName = "Service1";
        }

        private void onDataRecieved(object sender, SerialDataReceivedEventArgs e)
        {
            this.TakeLog("Data is Received");
            try
            {
                string str = this.serialPort.ReadExisting().ToString();
                this.TakeLog("Part of Data: " + str);
                if (str.StartsWith("\x0002") && !str.EndsWith("\x001a"))
                {
                    this.fullMSG = "";
                    this.fullMSG = this.fullMSG + str;
                }
                else if (!str.StartsWith("\x0002") && str.EndsWith("\x001a"))
                {
                    this.fullMSG = this.fullMSG + str;
                    this.TakeLog("Full RawData Received    ");
                    this.ExtractData(this.fullMSG);
                }
                else if (!str.StartsWith("\x0002") && !str.EndsWith("\x001a"))
                {
                    this.fullMSG = this.fullMSG + str;
                }
            }
            catch (Exception exception)
            {
                this.TakeLog("Exception Occured inside onDataReceived() Method: " + exception.Message.ToString());
            }
        }

        protected override void OnStart(string[] args)
        {
            try
            {
            }
            catch (Exception)
            {
                this.TakeLog("");
            }
            try
            {
                if (!Directory.Exists(logPath))
                {
                    Directory.CreateDirectory(logPath);
                }
            }
            catch (Exception exception2)
            {
                this.TakeLog("Exception Occured on Checking and Creating Log Path: " + exception2.Message.ToString());
            }
            try
            {
                this.con = new SqlConnection(this.Decryptor(File.ReadAllText(LisPath)));
            }
            catch (Exception exception3)
            {
                this.TakeLog("Exception Occured While Preparing SQL Connection. " + exception3.Message);
            }
            this.TakeLog("Service is Started!");
            this.TakeLog("The Connection string found is : " + File.ReadAllText(LisPath));
            try
            {
                this.OpenComPort();
            }
            catch (Exception exception4)
            {
                this.TakeLog("Exception Occured when Calling OpenComPort() Method: " + exception4.Message.ToString());
            }
        }

        protected override void OnStop()
        {
            if (this.serialPort.IsOpen)
            {
                this.serialPort.Close();
                this.TakeLog("Serial Port Successfully Close!");
            }
            this.TakeLog("Service Stoped!");
        }

        private void OpenComPort()
        {
            int num = -1;
            int num2 = -1;
            int num3 = -1;
            int num4 = -1;
            int num5 = -1;
            int num6 = -1;
            int num7 = -1;
            int num8 = -1;
            int num9 = -1;
            int num10 = -1;
            string str = File.ReadAllText(COMSettingPath);
            int num11 = 0;
            int num12 = 0;
            while (true)
            {
                if (num12 >= str.Length)
                {
                    string portName = str.Substring(num + 1, (num2 - num) - 1);
                    string str3 = str.Substring(num3 + 1, (num4 - num3) - 1);
                    string str4 = str.Substring(num5 + 1, (num6 - num5) - 1);
                    string str6 = str.Substring(num9 + 1, (num10 - num9) - 1);
                    StopBits stopBits = (str.Substring(num7 + 1, (num8 - num7) - 1) != "1") ? StopBits.Two : StopBits.One;
                    Parity parity = (str6 != "space") ? ((str6 != "even") ? ((str6 != "odd") ? ((str6 != "mark") ? Parity.None : Parity.Mark) : Parity.Odd) : Parity.Even) : Parity.Space;
                    try
                    {
                        this.serialPort = new SerialPort(portName, Convert.ToInt32(str3), parity, Convert.ToInt32(str4), stopBits);
                        this.serialPort.Handshake = Handshake.None;
                        this.serialPort.DataReceived += new SerialDataReceivedEventHandler(this.onDataRecieved);
                    }
                    catch (Exception exception)
                    {
                        this.TakeLog("error when opening com port: " + exception.Message);
                    }
                    break;
                }
                if (str[num12] == '|')
                {
                    num11++;
                    if (num11 == 1)
                    {
                        num = num12;
                    }
                    else if (num11 == 2)
                    {
                        num2 = num12;
                    }
                    else if (num11 == 3)
                    {
                        num3 = num12;
                    }
                    else if (num11 == 4)
                    {
                        num4 = num12;
                    }
                    else if (num11 == 5)
                    {
                        num5 = num12;
                    }
                    else if (num11 == 6)
                    {
                        num6 = num12;
                    }
                    else if (num11 == 7)
                    {
                        num7 = num12;
                    }
                    else if (num11 == 8)
                    {
                        num8 = num12;
                    }
                    else if (num11 == 9)
                    {
                        num9 = num12;
                    }
                    else if (num11 == 10)
                    {
                        num10 = num12;
                    }
                }
                num12++;
            }
            try
            {
                if (!this.serialPort.IsOpen)
                {
                    this.serialPort.Open();
                }
                string[] textArray1 = new string[11];
                textArray1[0] = "COM Port: ";
                textArray1[1] = this.serialPort.PortName;
                textArray1[2] = "   Baud Rate: ";
                textArray1[3] = this.serialPort.BaudRate.ToString();
                textArray1[4] = "   DataBits: ";
                textArray1[5] = this.serialPort.DataBits.ToString();
                textArray1[6] = "   StopBits: ";
                textArray1[7] = this.serialPort.StopBits.ToString();
                textArray1[8] = "   Parity: ";
                textArray1[9] = this.serialPort.Parity.ToString();
                textArray1[10] = "   ---Port Successfully Opened!";
                this.TakeLog(string.Concat(textArray1));
            }
            catch (Exception exception2)
            {
                string[] textArray2 = new string[12];
                textArray2[0] = "COM Port: ";
                textArray2[1] = this.serialPort.PortName;
                textArray2[2] = "   Baud Rate: ";
                textArray2[3] = this.serialPort.BaudRate.ToString();
                textArray2[4] = "   DataBits: ";
                textArray2[5] = this.serialPort.DataBits.ToString();
                textArray2[6] = "   StopBits: ";
                textArray2[7] = this.serialPort.StopBits.ToString();
                textArray2[8] = "   Parity: ";
                textArray2[9] = this.serialPort.Parity.ToString();
                textArray2[10] = "   --- Can Not Open COM Port! ";
                textArray2[11] = exception2.Message.ToString();
                this.TakeLog(string.Concat(textArray2));
            }
        }

        public void TakeLog(string text)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(new FileStream(logFilePath, FileMode.Append, FileAccess.Write)))
                {
                    writer.WriteLineAsync(DateTime.Now.ToString() + "\n\n  " + text + "   \n\n");
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}

