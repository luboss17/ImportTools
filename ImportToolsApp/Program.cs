using System;
using System.Data;
using System.Diagnostics;
using System.Net.Sockets;
using System.Threading;
using Docs.Excel;
using System.Drawing;

namespace ImportToolsApp
{
    
    class Program
    {
        private static TcpClient clientSocket = new TcpClient();
        private static Client pack;
        private const string IP = "";
        private const int major = 1;
        private const int minor = 0;
        private const int portNum = 58008;
        private const string insertCommand = "insert", updateCommand="update";
        private static DataTable table = new DataTable();
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            string toolsImportPath = @"C:\Cert Manager\ToolsList.xlsx";
            activateJetCellLicense();
            Console.Write("Excel File Path: ");
            string path = Console.ReadLine();
            path=toolsImportPath;
            if (path!="")
            {
                Console.WriteLine(path);
                importTools(path);
            }
         }
        private static void activateJetCellLicense()
        {
            const string jetCellLicense = "S3415N-781234-21BL6E-11AC00";
            ExcelWorkbook.SetLicenseCode(jetCellLicense);
        }
        private static void importTools(string path)
        {
            //run tool server
            runToolServer();
            //Logon to Tools Server as Guest_tool
            logon();
            //read file into table
            table=openTestExcel(path).Copy();
            //Write to toolCondTable
            DataTable toolsTable = writeToToolTable(table).Copy();
            //write to tool
            pack.insertMultipleTools(toolsTable);
            //close tool
        }
        //write passed in table data(raw data) to tooltable 
        private static DataTable writeToToolTable(DataTable oriTable)
        {
            const string command = "insert";
            pack.obj = "tool";
            //column index of where the fields are in oriTable
            const int toolCol = 0, modelCol = 2, manufactureCol = 1, snCol = 3, equipmentCol = 4;
            DataTable toolTable = pack.tools.Clone();
            try
            {
                for (int row=1;row<oriTable.Rows.Count;row++)
                {
                    DataRow newRow = toolTable.NewRow();
                    newRow[pack.toolID_colName] = oriTable.Rows[row][toolCol].ToString();
                    newRow[pack.model_colName] = oriTable.Rows[row][modelCol].ToString();
                    newRow[pack.equipment_colName] = oriTable.Rows[row][equipmentCol].ToString();
                    newRow[pack.manufacturer_colName] = oriTable.Rows[row][manufactureCol].ToString();
                    newRow[pack.SN_colName] = oriTable.Rows[row][snCol].ToString();
                    newRow[pack.scanOperator_colName] = true;
                    /*newRow[pack.lotID_colName] = lotID_txt.Text;
                    newRow[pack.scanOperator_colName] = convertBoolToInt(scanOperator_chk.Checked);
                    newRow[pack.scan1_colName] = convertBoolToInt(scan1_chk.Checked);
                    newRow[pack.scan2_colName] = convertBoolToInt(scan2_chk.Checked);
                    newRow[pack.setupPause_colName] = convertBoolToInt(pauseTool_chk.Checked);
                    newRow[pack.testID_colName] = testSetups_listBox.Text;
                    newRow[pack.mode_colName] = ch1Mode_comboBox.SelectedIndex;
                    newRow[pack.imode_colName] = ch2Mode_comboBox.SelectedIndex;*/

                    toolTable.Rows.Add(newRow);
                }   
            }
            catch (Exception e)
            {
                Console.WriteLine("Can't parse Excel to Tool Table due to error: " + e.Message);
            }
            return toolTable;
        }
        private static void runToolServer()
        {
            Process p = new Process();
            p.StartInfo = new ProcessStartInfo("cmd.exe", "/c cd C:\\Cert Manager & tools server");
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.CreateNoWindow = true;
            p.Start();
        }
        private static void connectServer()
        {
            //Connect to server
            if (clientSocket.Connected)
                clientSocket.Close();
            clientSocket = new System.Net.Sockets.TcpClient();
            try
            {
                tryToConnectToolDatabase();

            }
            catch (Exception e)
            {
                Thread.Sleep(3000);
                try
                {
                    tryToConnectToolDatabase();
                }
                catch
                {
                    Console.WriteLine("Fail to connect to Tools Database, please check to see if Tools server is installed\n"+e.Message);
                }
            }
        }
        private static void tryToConnectToolDatabase()
        {
            clientSocket.Connect(IP, portNum);
            if (clientSocket.Connected)//true if found server, false if server is not running
            {
                //serverMsgTxt.Text += "\nServer is connected";
                pack = new Client(clientSocket.GetStream(), major, minor);
            }
        }
        private static void logon()
        {
            connectServer();
            if (clientSocket.Connected)//if socket is connected with server, log on
            {
                pack.sendCommand("conn");
                pack.sendCommand("user");
            }
        }
        private static DataTable openTestExcel(string path)
        {
            ExcelWorkbook wbook = new ExcelWorkbook();
            if (path.EndsWith("xls"))
                wbook = ExcelWorkbook.ReadXLS(path);
            else if (path.EndsWith("xlsx"))
                wbook = ExcelWorkbook.ReadXLSX(path);
            else
            {
                Console.WriteLine("Invalid File Format, need to be xls or xlsx");
                return null;
            }
            DataTable table= wbook.Worksheets[0].WriteToDataTable();
            /*
            foreach (DataColumn column in table.Columns)
            {
                Console.WriteLine(column.ColumnName);
            }*/
            return table;
        }
    }
}
