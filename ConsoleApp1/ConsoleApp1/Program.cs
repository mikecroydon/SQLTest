using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Data;

namespace ConsoleApp8
{
    class Program
    {
        /*public void EditExcel(string myPath, int numBays, double dropRod, double lCant, double rCant) //Method to Edit an excel file
        {
            Application excel = new Application();  //Creates excel object an opens workbook

            Workbook workbook = excel.Workbooks.Open(myPath, ReadOnly: false, Editable: true);
            Worksheet worksheet = workbook.Worksheets.Item[1] as Worksheet;
            if (worksheet == null)
                return;

            string numBaysColString = "GE";         //Specifies the columns the different inputs are to be entered in
            string dropRodColString = "GF";
            string lCantColString = "GH";
            string rCantColString = "GI";

            int numBaysCol = ExcelColumnToNumber(numBaysColString);     //Converts columns strings to a number
            int dropRodCol = ExcelColumnToNumber(dropRodColString);
            int lCantCol = ExcelColumnToNumber(lCantColString);
            int rCantCol = ExcelColumnToNumber(rCantColString);

            Range numBaysLoc = worksheet.Rows.Cells[3, numBaysCol];     //Finds a location for different inputs using row and column
            Range dropRodLoc = worksheet.Rows.Cells[3, dropRodCol];
            Range lCantLoc = worksheet.Rows.Cells[3, lCantCol];
            Range rCantLoc = worksheet.Rows.Cells[3, rCantCol];

            numBaysLoc.Value = numBays;                //stores the inputs in the proper location
            dropRodLoc.Value = dropRod;
            lCantLoc.Value = lCant;
            rCantLoc.Value = rCant;

            excel.Application.ActiveWorkbook.Save();   //closes and saves workbook     
            excel.Application.Quit();
            excel.Quit();
        }

        public int ExcelColumnToNumber(string column)  //Method to convert a text string in excel format to a number. Double digit max
        {
            int columnNumber;
            string firstDigit, secondDigit;
            if (column.Length == 1)                    //Used for single digit strings
            {
                columnNumber = LetterToNumberString(column);
            }
            else                                       //Used for double digit strings
            {
                firstDigit = column.Substring(0, 1);   //Separates string into it's two pieces 
                secondDigit = column.Substring(1, 1);
                columnNumber = (LetterToNumberString(firstDigit) * 26) + (LetterToNumberString(secondDigit));
            }
            return columnNumber;
        }

        public int LetterToNumberString(string letter) //Switch to specify which letter corresponds to which number. A => 1, B=> 2 etc
        {
            switch (letter)
            {
                case "A":
                    return 1;
                    break;
                case "B":
                    return 2;
                    break;
                case "C":
                    return 3;
                    break;
                case "D":
                    return 4;
                    break;
                case "E":
                    return 5;
                    break;
                case "F":
                    return 6;
                    break;
                case "G":
                    return 7;
                    break;
                case "H":
                    return 8;
                    break;
                case "I":
                    return 9;
                    break;
                case "J":
                    return 10;
                    break;
                case "K":
                    return 11;
                    break;
                case "L":
                    return 12;
                    break;
                case "M":
                    return 13;
                    break;
                case "N":
                    return 14;
                    break;
                case "O":
                    return 15;
                    break;
                case "P":
                    return 16;
                    break;
                case "Q":
                    return 17;
                    break;
                case "R":
                    return 18;
                    break;
                case "S":
                    return 19;
                    break;
                case "T":
                    return 20;
                    break;
                case "U":
                    return 21;
                    break;
                case "V":
                    return 22;
                    break;
                case "W":
                    return 23;
                    break;
                case "X":
                    return 24;
                    break;
                case "Y":
                    return 25;
                    break;
                case "Z":
                    return 26;
                    break;
            }
            return 0;
        }
        static SqlDataReader myReader = null;

        public double ConvertDBToDouble(double CSName, string DBColumn)
        {

            if ((myReader[DBColumn]) != DBNull.Value) 
            {
                CSName = Convert.ToDouble(myReader[DBColumn]);
            }
            else{
                CSName = 0;
            }

            return CSName;
        }
        */

        public static void Main()
        {
            //variable declaration
            int numBays = 0;
            int dropRodL = 0;
            int numBridges = 0;
            int esc = 0;
            int b1Cap = 0;
            int DG1 = 0;
            int BD1 = 0;
            int TD1 = 0;
            int b2Cap = 0;
            int DG2 = 0;
            int BD2 = 0;
            int TD2 = 0;
            int b3Cap = 0;
            int DG3 = 0;
            int BD3 = 0;
            int TD3 = 0;
            int b4Cap = 0;
            int DG4 = 0;
            int BD4 = 0;
            int TD4 = 0;
            int b5Cap = 0;
            int DG5 = 0;
            int BD5 = 0;
            int TD5 = 0;
            int b6Cap = 0;
            int DG6 = 0;
            int BD6 = 0;
            int TD6 = 0;


            double runwayL = 0;
            double lCant = 0;
            double rCant = 0;
            double equalSCSpace = 0;
            double csc1 = 0;
            double csc2 = 0;
            double csc3 = 0;
            double csc4 = 0;
            double csc5 = 0;
            double csc6 = 0;
            double csc7 = 0;
            double csc8 = 0;
            double csc9 = 0;
            double csc10 = 0;
            double bridgeLen = 0;

            string sqlString;
            sqlString = SQL.ConnectionString(userID: "ASPNETUsers",
                password: "?Area51?",
                server: @"BION2DEV\HERONSQLDEV",
                database: "LiftLab_App");
            SQL.RetrieveCMTQuoteData(sqlString);

            //Passes values from SQL class properties defined in RetriveCMTQuoteData() to main() variables
            numBays = SQL.numBaysVal;
            dropRodL = SQL.dropRodLVal;
            numBridges = SQL.numBridgesVal;
            esc = SQL.escVal;
            b1Cap = SQL.b1CapVal;
            DG1 = SQL.DG1Val;
            BD1 = SQL.BD1Val;
            TD1 = SQL.TD1Val;
            b2Cap = SQL.b2CapVal;
            DG2 = SQL.DG2Val;
            BD2 = SQL.BD2Val;
            TD2 = SQL.TD2Val;
            b3Cap = SQL.b3CapVal;
            DG3 = SQL.DG3Val;
            BD3 = SQL.BD3Val;
            TD3 = SQL.TD3Val;
            b4Cap = SQL.b4CapVal;
            DG4 = SQL.DG4Val;
            BD4 = SQL.BD4Val;
            TD4 = SQL.TD4Val;
            b5Cap = SQL.b5CapVal;
            DG5 = SQL.DG5Val;
            BD5 = SQL.BD5Val;
            TD5 = SQL.TD5Val;
            b6Cap = SQL.b6CapVal;
            DG6 = SQL.DG6Val;
            BD6 = SQL.BD6Val;
            TD6 = SQL.TD6Val;

            runwayL = SQL.runwayLVal;
            lCant = SQL.lCantVal;
            rCant = SQL.rCantVal;
            equalSCSpace = SQL.equalSCSpaceVal;
            csc1 = SQL.csc1Val;
            csc2 = SQL.csc2Val;
            csc3 = SQL.csc3Val;
            csc4 = SQL.csc4Val;
            csc5 = SQL.csc5Val;
            csc6 = SQL.csc6Val;
            csc7 = SQL.csc7Val;
            csc8 = SQL.csc8Val;
            csc9 = SQL.csc9Val;
            csc10 = SQL.csc10Val;
            bridgeLen = SQL.bridgeLenVal;

            Console.WriteLine(numBays);
            Console.ReadLine();

            /*string quote = "CMT-123456";

            SqlConnection myConnection = new SqlConnection("user id=mcroydon;" +
                                                          "password=password;server=SPANCO-PC01;" +
                                                          "Trusted_Connection=yes;" +
                                                          "database=northwind; " +
                                                          "connection timeout=10");
            try
            {
                myConnection.Open();
            }
            catch(Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            int numBays = 0;
            int dropRodL = 0;
            int numBridges = 0;
            int esc = 0;
            int b1Cap = 0;
            int DG1 = 0;
            int BD1 = 0;
            int TD1 = 0;
            int b2Cap = 0;
            int DG2 = 0;
            int BD2 = 0;
            int TD2 = 0;
            int b3Cap = 0;
            int DG3 = 0;
            int BD3 = 0;
            int TD3 = 0;
            int b4Cap = 0;
            int DG4 = 0;
            int BD4 = 0;
            int TD4 = 0;
            int b5Cap = 0;
            int DG5 = 0;
            int BD5 = 0;
            int TD5 = 0;
            int b6Cap = 0;
            int DG6 = 0;
            int BD6 = 0;
            int TD6 = 0;

            
            double runwayL = 0;
            double lCant = 0;
            double rCant = 0;
            double equalSCSpace = 0;
            double csc1 = 0;
            double csc2 = 0;
            double csc3 = 0;
            double csc4 = 0;
            double csc5 = 0;
            double csc6 = 0;
            double csc7 = 0;
            double csc8 = 0;
            double csc9 = 0;
            double csc10 = 0;
            double bridgeLen = 0;

            SqlCommand myCommand = new SqlCommand("SELECT * FROM LiftLabQuotes WHERE QuoteID=@QuoteID", myConnection);

            SqlParameter quoteID = new SqlParameter();
            quoteID.ParameterName = "@QuoteId";
            quoteID.Value = quote;

            myCommand.Parameters.Add(quoteID);

            myReader = myCommand.ExecuteReader();

            while (myReader.Read())
            {


                numBays = Convert.ToInt32(myReader["NumBays"]);
                dropRodL = Convert.ToInt32(myReader["DropRodL"]);
                numBridges = Convert.ToInt32(myReader["NumBridges"]);
                esc = Convert.ToInt32(myReader["EqualSC"]);
                b1Cap = Convert.ToInt32(myReader["Bridge1Cap"]);
                DG1 = Convert.ToInt32(myReader["Bridge1DG"]);
                BD1 = Convert.ToInt32(myReader["Bridge1Drive"]);
                TD1 = Convert.ToInt32(myReader["Bridge1TrolleyDrive"]);
                b2Cap = Convert.ToInt32(myReader["Bridge1Cap"]);
                DG2 = Convert.ToInt32(myReader["Bridge1DG"]);
                BD2 = Convert.ToInt32(myReader["Bridge1Drive"]);
                TD2 = Convert.ToInt32(myReader["Bridge1TrolleyDrive"]);
                b3Cap = Convert.ToInt32(myReader["Bridge1Cap"]);
                DG3 = Convert.ToInt32(myReader["Bridge1DG"]);
                BD3 = Convert.ToInt32(myReader["Bridge1Drive"]);
                TD3 = Convert.ToInt32(myReader["Bridge1TrolleyDrive"]);
                b4Cap = Convert.ToInt32(myReader["Bridge1Cap"]);
                DG4 = Convert.ToInt32(myReader["Bridge1DG"]);
                BD4 = Convert.ToInt32(myReader["Bridge1Drive"]);
                TD4 = Convert.ToInt32(myReader["Bridge1TrolleyDrive"]);
                b5Cap = Convert.ToInt32(myReader["Bridge1Cap"]);
                DG5 = Convert.ToInt32(myReader["Bridge1DG"]);
                BD5 = Convert.ToInt32(myReader["Bridge1Drive"]);
                TD5 = Convert.ToInt32(myReader["Bridge1TrolleyDrive"]);
                b6Cap = Convert.ToInt32(myReader["Bridge1Cap"]);
                DG6 = Convert.ToInt32(myReader["Bridge1DG"]);
                BD6 = Convert.ToInt32(myReader["Bridge1Drive"]);
                TD6 = Convert.ToInt32(myReader["Bridge1TrolleyDrive"]);

                Program convert = new Program();
                
                runwayL = convert.ConvertDBToDouble(runwayL, "RunwayL");
                lCant = convert.ConvertDBToDouble(lCant, "LCant");
                rCant = convert.ConvertDBToDouble(rCant, "RCant");
                equalSCSpace = convert.ConvertDBToDouble(equalSCSpace, "EqualSCSpacing");
                csc1 = convert.ConvertDBToDouble(csc1, "CSC1");
                csc2 = convert.ConvertDBToDouble(csc2, "CSC2");
                csc3 = convert.ConvertDBToDouble(csc3, "CSC3");
                csc4 = convert.ConvertDBToDouble(csc4, "CSC4");
                csc5 = convert.ConvertDBToDouble(csc5, "CSC5");
                csc6 = convert.ConvertDBToDouble(csc6, "CSC6");
                csc7 = convert.ConvertDBToDouble(csc7, "CSC7");
                csc8 = convert.ConvertDBToDouble(csc8, "CSC8");
                csc9 = convert.ConvertDBToDouble(csc9, "CSC9");
                csc10 = convert.ConvertDBToDouble(csc10, "CSC10");
                bridgeLen = convert.ConvertDBToDouble(bridgeLen, "BridgeLength");
              
            }

            Console.WriteLine(b1Cap);
            Console.WriteLine(bridgeLen);
            Console.ReadLine();

            //Console.WriteLine(RunwayL).ToString());
            /*try
            {
                SqlDataReader myReader = null;
                SqlCommand myCommand = new SqlCommand("select * from LiftLabQuotes where quoteID = 'CMT-123456'",
                                                         myConnection);
                myReader = myCommand.ExecuteReader();
                while (myReader.Read())

                {
                    Console.WriteLine(myReader["RunwayL"].ToString());
                    Console.WriteLine(myReader["LCant"].ToString());
                }
             
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            Thread.Sleep(10000);

            try
            {
                myConnection.Close();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            Program WriteExcel = new Program();

             string myPath = @"C:\Users\mcroydon\Documents\920.xlsx";
             int numBays = 7;
             int dropRod = 2;
             int lCant = 3;
             int rCant = 3;

             WriteExcel.EditExcel(myPath, numBays, dropRod, lCant, rCant);
             */
        }

    }
}