using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Sockets;
using System.Data;
using System.Security.Cryptography;
using System.Threading;

namespace ImportToolsApp
{
    class Client
    {
        public NetworkStream serverstream;

        private int major, minor;
        private byte[] outBuffer = new byte[65536];
        public byte[] inBuffer = new byte[65536];
        public string obj = "";
        private string clientName = "Client v1.0.0";

        public DataTable tools = new DataTable(), result = new DataTable(), sample = new DataTable(), last = new DataTable(), error = new DataTable(), conditions = new DataTable(), resultTable;
        public bool resultUpdate = false;
        private byte[] keyBytes = new byte[32]; //Key blob received from server after connected to server. Used to login
        public Client(NetworkStream networkstream, int major, int minor)
        {
            this.serverstream = networkstream;
            this.major = major;
            this.minor = minor;

            defineTable();
        }
        //return result Table
        public DataTable getResultTable()
        {
            //DataRow row = last.Rows[last.Rows.Count-1];

            return resultTable;
        }
        /// <summary>
        /// add column to table, type: 1=string,2=int,3=double,4=Int64
        /// </summary>
        /// <param name="table"></param>
        /// <param name="columnName"></param>
        /// <param name="type"></param>
        private void addColumn(ref DataTable table, string columnName, int type)
        {
            string getType = "";
            DataColumn col = new DataColumn(columnName);
            switch (type)
            {
                case 1:
                    getType = "System.String";
                    break;
                case 2:
                    getType = "System.Int32";
                    break;
                case 3:
                    getType = "System.Double";
                    break;
                case 4:
                    getType = "System.Int64";
                    break;
            }
            col.DataType = System.Type.GetType(getType);
            table.Columns.Add(col);
        }

        //define name of all dataTable colName
        public string toolID_colName = "tool_id", model_colName = "model", equipment_colName = "equipment", manufacturer_colName = "manufacturer", SN_colName = "serial_num", lotID_colName = "lot_id", input_colName = "input",
            iunitIndex_colName = "iunits index", imin_colName = "imin", imax_colName = "imax", imode_colName = "imode", output_colName = "output", ounitIndex_colName = "ounits index", omin_colName = "omin", omax_colName = "omax", mode_colName = "mode", freq_colName = "freq",
            testID_colName = "test_id", procedure_colName = "procedure", pointIndex_colName = "points index", sign_colName = "sign", testOn_colName = "test_on", useToolFlag_colName = "use_tool_flag", tmax_colName = "tmax",
            tunitIndex_colName = "tunits index", sunits_colName = "sunits", count_colName = "count", sauto_colName = "sauto", lowTolerance_colName = "low tolerance", highTolerance_colName = "high tolerance",
            tolType_colName = "tol_type", criteria_colName = "criteria", tauto_colName = "tauto", stopfail_colName = "stop_fail", scanOperator_colName = "scan_operator flag", scan1_colName = "scan_comment1 flag",
            scan2_colName = "scan_comment2 flag", setupPause_colName = "setup_pause flag";
        //set up column for each object tools,result, sample. These tables are used to store result from "find" response
        private void defineTable()
        {
            addColumn(ref tools, toolID_colName, 1);
            addColumn(ref tools, model_colName, 1);
            addColumn(ref tools, equipment_colName, 1);
            addColumn(ref tools, manufacturer_colName, 1);
            addColumn(ref tools, SN_colName, 1);
            addColumn(ref tools, lotID_colName, 1);
            addColumn(ref tools, input_colName, 2);
            addColumn(ref tools, iunitIndex_colName, 2);//verify if the name for column is right
            addColumn(ref tools, imin_colName, 3);
            addColumn(ref tools, imax_colName, 3);
            addColumn(ref tools, imode_colName, 2);
            addColumn(ref tools, output_colName, 2);
            addColumn(ref tools, ounitIndex_colName, 2);
            addColumn(ref tools, omin_colName, 3);
            addColumn(ref tools, omax_colName, 3);
            addColumn(ref tools, mode_colName, 2);
            addColumn(ref tools, freq_colName, 2);
            addColumn(ref tools, testID_colName, 1);
            addColumn(ref tools, procedure_colName, 1);
            addColumn(ref tools, pointIndex_colName, 2);
            addColumn(ref tools, sign_colName, 2);
            addColumn(ref tools, testOn_colName, 2);
            addColumn(ref tools, useToolFlag_colName, 2);
            addColumn(ref tools, tmax_colName, 3);
            addColumn(ref tools, tunitIndex_colName, 2);
            addColumn(ref tools, sunits_colName, 2);
            addColumn(ref tools, count_colName, 2);
            addColumn(ref tools, sauto_colName, 2);
            addColumn(ref tools, lowTolerance_colName, 3);
            addColumn(ref tools, highTolerance_colName, 3);
            addColumn(ref tools, tolType_colName, 2);
            addColumn(ref tools, criteria_colName, 2);
            addColumn(ref tools, tauto_colName, 2);
            addColumn(ref tools, stopfail_colName, 2);
            addColumn(ref tools, scanOperator_colName, 2);
            addColumn(ref tools, scan1_colName, 2);
            addColumn(ref tools, scan2_colName, 2);
            addColumn(ref tools, setupPause_colName, 2);

            addColumn(ref result, "id", 4);
            addColumn(ref result, "operator", 1);
            addColumn(ref result, "comment1", 1);
            addColumn(ref result, "comment2", 1);
            addColumn(ref result, "label1 for comment 1", 1);
            addColumn(ref result, "label2 for comment 2", 1);
            addColumn(ref result, "timestamp", 1);
            addColumn(ref result, "input_sn", 1);
            addColumn(ref result, "output_sn", 1);
            addColumn(ref result, "customer", 1);
            addColumn(ref result, "po", 1);

            addColumn(ref result, "tool_id", 1);
            addColumn(ref result, "model", 1);
            addColumn(ref result, "equipment", 1);
            addColumn(ref result, "manufacturer", 1);
            addColumn(ref result, "serial_num", 1);
            addColumn(ref result, "lot_id", 1);
            addColumn(ref result, "input", 2);
            addColumn(ref result, "iunits index", 2);//verify if the name for column is right
            addColumn(ref result, "imin", 3);
            addColumn(ref result, "imax", 3);
            addColumn(ref result, "imode", 2);
            addColumn(ref result, "output", 2);
            addColumn(ref result, "ounits index", 2);
            addColumn(ref result, "omin", 3);
            addColumn(ref result, "omax", 3);
            addColumn(ref result, "mode", 2);
            addColumn(ref result, "freq", 2);
            addColumn(ref result, "test_id", 1);
            addColumn(ref result, "procedure", 1);
            addColumn(ref result, "points index", 2);
            addColumn(ref result, "sign", 2);
            addColumn(ref result, "test_on", 2);
            addColumn(ref result, "use_tool flag", 2);
            addColumn(ref result, "tmax", 3);
            addColumn(ref result, "tunits index", 2);
            addColumn(ref result, "sunits", 2);
            addColumn(ref result, "count", 2);
            addColumn(ref result, "sauto", 2);
            addColumn(ref result, "low tolerance", 3);
            addColumn(ref result, "high tolerance", 3);
            addColumn(ref result, "tol_type", 2);
            addColumn(ref result, "criteria", 2);
            addColumn(ref result, "tauto", 2);
            addColumn(ref result, "stop_fail", 2);
            addColumn(ref result, "scan_operator flag", 2);
            addColumn(ref result, "scan_comment1 flag", 2);
            addColumn(ref result, "scan_comment2 flag", 2);
            addColumn(ref result, "setup_pause flag", 2);

            addColumn(ref sample, "id", 4);
            addColumn(ref sample, "result id", 4);
            addColumn(ref sample, "input sample", 1);
            addColumn(ref sample, "input units", 1);
            addColumn(ref sample, "output sample", 1);
            addColumn(ref sample, "output units", 1);
            addColumn(ref sample, "target", 1);
            addColumn(ref sample, "low limit", 1);
            addColumn(ref sample, "high limit", 1);
            addColumn(ref sample, "result", 1);
            addColumn(ref sample, "timestamp", 1);

            addColumn(ref last, "response type", 1);
            addColumn(ref last, "response count", 2);

            addColumn(ref error, "description of the error", 1);
        }

        //write length of buffer to first 3 bytes of buffer[]
        private byte[] writeLen(byte[] buffer, int len)
        {
            buffer[0] = Convert.ToByte((len & 0x3F) + '0');
            buffer[1] = Convert.ToByte(((len >> 6) & 0x3F) + '0');
            buffer[2] = Convert.ToByte(((len >> 12) & 0x3F) + '0');
            return buffer;
        }
        //read the length of buffer from first 3 bytes of buffer[]
        //this method is used by receiveCommand to decode message received by server
        private int readLen(byte[] buffer)
        {
            int len = 0;
            if (buffer[3] != (byte)0)
                return 0;
            else
            {
                len += buffer[0] - '0';
                len += (buffer[1] - '0') << 6;
                len += (buffer[2] - '0') << 12;
                return len;
            }
        }

        //Pass in string, position to write string on buffer[]
        private byte[] writeStrtoByte(string str, int start, byte[] buffer)
        {
            byte[] byteToWrite = new byte[str.Length];
            byteToWrite = Encoding.UTF8.GetBytes(str);
            for (int i = 0; i < str.Length; i++)
            {
                buffer[i + start] = byteToWrite[i];
            }
            return buffer;
        }
        ///param: start, end position of byte[] buffer, return string decoded from buffer[start -> end]
        private string readBytetoStr(int start, int end, byte[] buffer)
        {
            byte[] byteToRead = new byte[end - start + 1];
            string returnStr = "";
            int index = 0;
            for (int i = start; i <= end; i++)
            {
                byteToRead[index] = buffer[i];
                index++;
            }
            returnStr = Encoding.ASCII.GetString(byteToRead);
            return returnStr;
        }
        private int nextNullLocation(byte[] buffer, int start)
        {
            int end = start;
            while (buffer[end] != (byte)0)
            {
                end++;
            }
            return end;
        }

        //compute and return SHA256 format password to send to tools server, Param: username, password and keyBytes from server respond to conn command
        private byte[] computePW_SHA256(string userName, string password, byte[] keyBytes)
        {

            SHA256 pwSha256 = SHA256Managed.Create();
            int userANDpwByteLen = userName.Length + password.Length + 2;//userANDpwbyteLen is equal username and password length + 2 extra null byte
            byte[] pwANDuserBytes = new byte[userANDpwByteLen];
            pwANDuserBytes = writeStrtoByte(password, 0, pwANDuserBytes);
            pwANDuserBytes = writeStrtoByte(userName, password.Length + 1, pwANDuserBytes);
            string temp = System.Text.Encoding.UTF8.GetString(pwANDuserBytes);
            //Compute SHA256 of username and password
            byte[] db_hash;
            db_hash = pwSha256.ComputeHash(pwANDuserBytes);

            //Copy both db_hash and keyBytes to create final password_hash
            byte[] preSHA_pwByte = new byte[db_hash.Length + keyBytes.Length];//preSHA_pwByte will contain both db_hash bytes and keyBytes, plus additional null byte in between
            db_hash.CopyTo(preSHA_pwByte, 0);
            keyBytes.CopyTo(preSHA_pwByte, db_hash.Length);

            //Compute final password hash to send to server
            byte[] password_hash = pwSha256.ComputeHash(preSHA_pwByte);

            return password_hash;
        }
        private const string deleteKey = "DELETE";
        private const string userName = "tools_guest", password = "torque\r123";//User name and password to log into server. Todo: update these when known

        public string toolIDToDelete = "";
        //Main Method used to send command to server
        public void sendCommand(string command)
        {
            int len = 0;
            outBuffer = new byte[65536];
            int countingPos = 0; //Running possition of where we are writing to outBuffer
            countingPos += 4;//Add 3 bytes for lenghth and 1 null byte
            Random rnd = new Random();
            string sequenceNum = rnd.Next(1, 9).ToString();
            //Write command to outbuffer    
            outBuffer = writeStrtoByte(command, countingPos, outBuffer); //write command to outBuffer
            countingPos += command.Length + 1; //Add length of command and extra null byte
            string conditionsCount;
            switch (command)
            {
                case "conn"://attablish connection with server. Need to be the first thing to be called
                    {

                        outBuffer = writeStrtoByte(major.ToString(), countingPos, outBuffer);
                        countingPos += major.ToString().Length + 1;
                        outBuffer = writeStrtoByte(minor.ToString(), countingPos, outBuffer);
                        countingPos += minor.ToString().Length + 1;
                        outBuffer = writeStrtoByte(clientName, countingPos, outBuffer);
                        countingPos += clientName.Length + 1;

                        len = countingPos;
                        outBuffer = writeLen(outBuffer, len);//write length to first 3 bytes
                        break;
                    }
                case "user"://logon to enable modifying database
                    {
                        //Write username
                        outBuffer = writeStrtoByte(userName, countingPos, outBuffer);
                        countingPos += userName.Length + 1;

                        //Write password
                        byte[] pwBytes = new byte[33];

                        //Write length of blob32 (32) to next byte
                        pwBytes = writeStrtoByte(" ", 0, pwBytes);//write 32 to first byte

                        //Compute pw_hashBytes and write to next 32 bytes in outbuffer
                        byte[] pw_hashBytes = computePW_SHA256(userName, password, keyBytes);
                        Array.Copy(pw_hashBytes, 0, pwBytes, 1, 32);//write content of SHA256 pw to next 32 byte
                        pwBytes.CopyTo(outBuffer, countingPos);//copy pwBytes to outbuffer 
                        countingPos += pwBytes.Length + 1;
                        len = countingPos;

                        //Write length of outbuffer
                        outBuffer = writeLen(outBuffer, len);
                        break;
                    }
                case "find"://Note: obj is tools,results,samples,points
                    {
                        //Write Object
                        outBuffer = writeStrtoByte(obj, countingPos, outBuffer);//write object: tools/results,samples
                        countingPos += obj.Length + 1;

                        //Write condition count to outbuffer
                        conditionsCount = conditions.Columns.Count.ToString();
                        outBuffer = writeStrtoByte(conditionsCount, countingPos, outBuffer);//write count: how many conditionss
                        countingPos += conditionsCount.Length + 1;

                        //Write the conditions
                        for (int i = 0; i < conditions.Columns.Count; i++)
                        {
                            string colName = conditions.Columns[i].ColumnName;
                            string value = conditions.Rows[0][i].ToString();
                            string strToWrite = colName + "?" + value + "%";//strToWrite have this format example:tool_id?qc%
                            outBuffer = writeStrtoByte(strToWrite, countingPos, outBuffer);
                            countingPos += strToWrite.Length + 1;
                        }

                        //Write Length
                        len = countingPos;
                        outBuffer = writeLen(outBuffer, len);
                        break;
                    }
                case "delete"://Note: obj is tool,result,sample,point
                    {
                        if (toolIDToDelete != "")
                        {
                            //Write sequence number

                            outBuffer = writeStrtoByte(sequenceNum, countingPos, outBuffer);
                            countingPos += sequenceNum.Length + 1;

                            //Write DELETE key
                            outBuffer = writeStrtoByte(deleteKey, countingPos, outBuffer);
                            countingPos += deleteKey.Length + 1;

                            //Write object type to outbuffer
                            outBuffer = writeStrtoByte(obj, countingPos, outBuffer);
                            countingPos += obj.Length + 1;

                            //Write conditions to buffer
                            outBuffer = writeStrtoByte(toolIDToDelete, countingPos, outBuffer);
                            countingPos += toolIDToDelete.Length + 1;

                            //Write length of a whole buffer to first 3 bytes
                            outBuffer = writeLen(outBuffer, countingPos);

                            len = countingPos;
                        }
                        break;
                    }
                case "update"://Note: obj is tool,result,sample,point
                    //Write sequence number
                    outBuffer = writeStrtoByte(sequenceNum, countingPos, outBuffer);
                    countingPos += sequenceNum.Length + 1;

                    //Write object type to outbuffer
                    outBuffer = writeStrtoByte(obj, countingPos, outBuffer);
                    countingPos += obj.Length + 1;

                    //Write the conditions
                    foreach (DataRow dataRow in conditions.Rows)
                    {
                        for (int colIndex = 0; colIndex < conditions.Columns.Count; colIndex++)
                        {
                            if (dataRow[colIndex] != null)
                            {
                                string strToWrite = dataRow[colIndex].ToString();
                                outBuffer = writeStrtoByte(strToWrite, countingPos, outBuffer);
                                countingPos += strToWrite.Length + 1;
                            }
                            else
                            {
                                countingPos++;
                            }
                        }
                    }

                    //Write Length
                    len = countingPos;
                    outBuffer = writeLen(outBuffer, len);

                    break;
                case "insert"://Note: obj is tool,result,sample,point
                    //Write sequence number
                    outBuffer = writeStrtoByte(sequenceNum, countingPos, outBuffer);
                    countingPos += sequenceNum.Length + 1;

                    //Write Object
                    outBuffer = writeStrtoByte(obj, countingPos, outBuffer);//write object: tools/results,samples
                    countingPos += obj.Length + 1;

                    //Write the conditions
                    foreach (DataRow dataRow in conditions.Rows)
                    {
                        for (int colIndex = 0; colIndex < conditions.Columns.Count; colIndex++)
                        {
                            if (dataRow[colIndex] != null)
                            {
                                string strToWrite = dataRow[colIndex].ToString();
                                outBuffer = writeStrtoByte(strToWrite, countingPos, outBuffer);
                                countingPos += strToWrite.Length + 1;
                            }
                            else
                            {
                                countingPos++;
                            }
                        }
                    }

                    //Write Length
                    len = countingPos;
                    outBuffer = writeLen(outBuffer, len);
                    break;
            }

            //write outbuffer to server
            serverstream.Write(outBuffer, 0, len);

            //Sleep half a sec to wait for serv to respond
            Thread.Sleep(500);

            resultTable = new DataTable();
            receiveServerResopond(command, sequenceNum);
            serverstream.Flush();
            resultUpdate = true;
        }
        public void insertMultipleTools(DataTable toolsTable)
        {
            foreach (DataRow dataRow in toolsTable.Rows)
            {
                string command = "insert";
                int len = 0;
                outBuffer = new byte[65536];
                int countingPos = 0; //Running possition of where we are writing to outBuffer
                countingPos += 4;//Add 3 bytes for lenghth and 1 null byte
                Random rnd = new Random();
                string sequenceNum = rnd.Next(1, 9).ToString();
                //Write command to outbuffer    
                outBuffer = writeStrtoByte(command, countingPos, outBuffer); //write command to outBuffer
                countingPos += command.Length + 1; //Add length of command and extra null byte

                //Write sequence number
                outBuffer = writeStrtoByte(sequenceNum, countingPos, outBuffer);
                countingPos += sequenceNum.Length + 1;

                //Write Object
                outBuffer = writeStrtoByte(obj, countingPos, outBuffer);//write object: tools/results,samples
                countingPos += obj.Length + 1;

                //Write the conditions

                for (int colIndex = 0; colIndex < toolsTable.Columns.Count; colIndex++)
                {
                    if (dataRow[colIndex] != null)
                    {
                        string strToWrite = dataRow[colIndex].ToString();
                        outBuffer = writeStrtoByte(strToWrite, countingPos, outBuffer);
                        countingPos += strToWrite.Length + 1;
                    }
                    else
                    {
                        countingPos++;
                    }
                }


                //Write Length
                len = countingPos;
                outBuffer = writeLen(outBuffer, len);

                //write outbuffer to server
                serverstream.Write(outBuffer, 0, len);

                //Sleep half a sec to wait for serv to respond
                Thread.Sleep(500);

                resultTable = new DataTable();
                receiveServerResopond(command, sequenceNum);
                serverstream.Flush();
                resultUpdate = true;
            }
            Console.WriteLine("updated");
        }
        //Call this method to read from server, pass in serverstream,buffer to write to,offset and length of message
        private byte[] readServer(NetworkStream serverstream, byte[] buffer, int offset, int len)
        {

            if (serverstream.DataAvailable)
                serverstream.Read(buffer, offset, len);
            return buffer;
        }
        //assign buffer to appropriate table. Pass in buffer, start byte location,length of a whole buffer and the table. Called by decodeInBuffer
        private DataTable dataToTable(byte[] buffer, int start, int end, DataTable table)
        {
            DataRow row = table.NewRow();
            int colCount = 0;
            while (start < end)
            {
                int beforeNul = nextNullLocation(buffer, start) - 1;
                row[colCount] = readBytetoStr(start, beforeNul, buffer);
                colCount++;
                start = beforeNul + 2;
            }
            table.Rows.Add(row);
            return table;
        }

        //pass in byte[] from server, decode and put in appropriate table
        private void decodeInBuffer(byte[] buffer, int len)
        {
            string command = readBytetoStr(4, 7, buffer);//read in type of command, should be "find"
            int count = 0;
            string obj = "";
            if (command == "find")
            {
                count = Convert.ToInt32(readBytetoStr(9, 9, buffer));//get how many results server sent back
                int objNullLoc = nextNullLocation(buffer, 11);
                obj = readBytetoStr(11, objNullLoc - 1, buffer);//get object of this packet, can either be tools,server,error, last etc.
                if (resultTable.Rows.Count < 1)//if this is the first response from server, assign header for resultTable
                {
                    switch (obj)
                    {
                        case "tool":
                            resultTable = tools.Clone();
                            break;
                        case "result":
                            resultTable = result.Clone();
                            break;
                        case "sample":
                            resultTable = sample.Clone();
                            break;
                        case "last"://read remaining byte to buffer if there is any other data available to be read.
                            buffer = readServer(serverstream, buffer, 0, buffer.Length);
                            break;
                        case "error":
                            resultTable = error.Clone();
                            break;
                    }
                }
                if (obj != "last")
                    resultTable = dataToTable(buffer, objNullLoc + 1, len - 1, resultTable);
            }

        }

        //Main Method used to receive command from server in byte[] format
        private void receiveServerResopond(string command, string sequenceNum)
        {
            byte[] buffer = new Byte[65536];
            int offset = 0;
            buffer = readServer(serverstream, buffer, offset, 4);//Read in first 4 byte that contains length of receiving message first
            offset += 4;
            int len = readLen(buffer);
            string thisObj = "";

            if (len > 0)
            {
                //read everything into buffer
                buffer = readServer(serverstream, buffer, offset, len - offset);
                switch (command)
                {
                    case "conn":
                        //Todo: Handle when the flag is set to password reset
                        Array.Copy(buffer, len - 32, keyBytes, 0, 32);//copy the last 32 bytes = blob key into keyBytes
                        break;
                    case "find":
                        //read object into thisObject
                        int objNullLoc = nextNullLocation(buffer, 11);
                        thisObj = readBytetoStr(11, objNullLoc - 1, buffer);

                        decodeInBuffer(buffer, len);//decode message received and put them in appropriate table

                        //if this is not the last line of respond, keep reading next line
                        if (thisObj != "last")
                        {
                            receiveServerResopond(command, sequenceNum);
                        }
                        break;
                    case "delete":
                        //Todo: Use response to see whether delete succeeded

                        break;
                    case "insert":
                        //Todo: Use response to see whether insert succeed

                        break;
                    case "update":
                        //Todo: Use response to see whether insert succeed
                        break;

                }

                //Thread newThread = new Thread(() => decodeInBuffer(buffer));
                //newThread.Start();
            }
            else
            {
                buffer = new Byte[65536];
                serverstream.Flush();
            }

            //if command is find and this is not the last message, continute reading from server


            //return buffer;
            //add code to add row to table instead
        }
    }
}
