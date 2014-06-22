using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data.SqlClient;


namespace TechInventoryReport
{
   
 
    class reportdata
    {
        private string mOracleid;
        private string mDeptName;
        public enum myCatetories {There, Missing, Never, Elsewhere, EOLSeen}
        private System.Collections.Hashtable title;
        private System.Collections.Hashtable fields;
        private System.Collections.Hashtable description;
        private int mLogID;
        private DateTime mReportDate;
        private string[] sortby = { "Room,ASSET_TAG", "Room,ASSET_TAG", "MODEL,Asset_tag", "Location,ASSET_TAG", "Room,ASSET_TAG" };

        internal String OracleID
        {
            get
            {
                return mOracleid;
            }
            set
            {
                mOracleid = value;
            }
        }

        internal int LogID
        {
            get
            {
                return mLogID;
            }
            set
            {
                mLogID = value;
            }

        }
        internal DateTime reportDate
        {    
            get
            {
                return mReportDate;
            }
            set
            {
                mReportDate = value;
            }


        }


        internal String deptName
        {
            get
            {
                return mDeptName;
            }
            set
            {
                mDeptName = value;
            }
        }

        internal reportdata(string dept, string date, iTextSharp.text.Document doc, int LogID)  //constructor
        {
            this.OracleID = dept;
            this.LogID = LogID;
            

            Console.WriteLine("Oracle ID " + OracleID);
             title=new System.Collections.Hashtable();
             fields = new System.Collections.Hashtable();
             description= new System.Collections.Hashtable();


             description.Add("There", "Part I is a list of technology assets that are “in good standing” in your school – as in, these assets are in the Fixed Assets Application (FAA) system (“the inventory”) and they have been seen on the CPS network in the last 30 days.");
             description.Add("Missing", "Part II is a list of technology assets that have been seen on the CPS network at some point in time, but have not been seen on the network in the last 30 days.");
             description.Add("Never", "Part III is a list of technology assets that are in your school’s inventory, but that have never been seen on the CPS network.");
             description.Add("Elsewhere", "Part IV is a list of technology assets that are in your school’s inventory, but that have recently been seen elsewhere on the CPS network.");
             description.Add("EOLSeen", "Part V is a list of technology assets that have been stolen or disposed according to your school’s inventory, but that have been seen on the network since the date stolen/disposed. ");
            


            title.Add("There", "Part I: What’s There");
            title.Add("Missing", "Part II: What’s Missing");
            title.Add("Never", "Part III: Never Seen");
            title.Add("Elsewhere", "Part IV: Seen Elsewhere");
            title.Add("EOLSeen", "Part V: Marked as Stolen/Disposed but still seen");

           
            string[] seenFields={"Model", "Serial Number", "Asset Tag", "Room"};
            string[] neverFields= {"Model", "Serial Number", "Asset Tag", "Room","Asset ID", "Description"};
            string[] elswewhereFields = {"Description", "Serial Number", "Asset Tag", "Asset ID", "Location" };
            string[] EOLseenFields = { "Asset ID", "Description", "Serial Number", "Asset Tag", "Room" };

            fields.Add("1",seenFields);
            fields.Add("2", seenFields); //not seen lately has the same fields 
            fields.Add("3",neverFields);
            fields.Add("4",elswewhereFields);
            fields.Add("5",EOLseenFields);


        }



        internal bool GetSummaryInfo(ref string deptName, ref int numThere, ref int numMissing, ref int numNever, ref int numElsewhere, ref int numEOLSeen)
        {
            SqlConnection cnn = dbConnect();
            if (cnn == null)
            {
                return false;
            }
            else
            {
                try
                {
                    StringBuilder sbCmd = new StringBuilder();
                    sbCmd.Append("select h.*, MS_SHORTNM, DEPARTMENT_NAME from TECH_INVENTORY_HISTORY h ");
                    sbCmd.Append("inner join (select max(run_date) as run_date, oracle_id from TECH_INVENTORY_HISTORY group by oracle_id) g ");
                    sbCmd.Append("on g.oracle_id=h.oracle_id  and g.run_date=h.RUN_DATE ");
                    sbCmd.Append("inner join TECHXL_DEPARTMENTS d on d.ORACLE_ID=h.ORACLE_ID ");                    
                    sbCmd.Append("where h.oracle_id=@Oracle_ID and h.run_date=g.run_date ");
                    sbCmd.Append("and exists (select * from  TEMP_TECH_INVENTORY t where t.PARENT_LOG_ID=h.TECH_INVENTORY_LOG_ID)");

                    SqlCommand thisCommand = new SqlCommand(sbCmd.ToString(), cnn);                   

                    thisCommand.CommandTimeout = 120;
                    thisCommand.Parameters.Add("@Oracle_ID", System.Data.SqlDbType.VarChar);
                    thisCommand.Parameters["@Oracle_ID"].Value = OracleID;
                    SqlDataReader thisReader = thisCommand.ExecuteReader();
                    int i;

                    if (thisReader.HasRows)
                    {
                        thisReader.Read();
                        deptName = thisReader["DEPARTMENT_NAME"].ToString();
                        DateTime D;
                        if (DateTime.TryParse(thisReader["RUN_DATE"].ToString(), out D))
                        {
                            this.reportDate=D;
                         } else {
                             this.reportDate=DateTime.Now;
                         }
                        if (int.TryParse(thisReader["TECH_INVENTORY_LOG_ID"].ToString(), out i))
                        {

                            //data exist for this Oracle ID
                            this.LogID = i;
                            numThere=int.Parse(thisReader["SEEN"].ToString());
                            numMissing=int.Parse(thisReader["MISSING"].ToString());
                            numNever = int.Parse(thisReader["NEVER_SEEN"].ToString());
                            numElsewhere = int.Parse(thisReader["SEEN_ELSEWHERE"].ToString());
                            numEOLSeen = int.Parse(thisReader["EOL_SEEN"].ToString());
                         
                        }
                        else
                        {
                            return false;

                        }
                    }


                    return true;
                }
                catch
                {
                    return false;
                }
            }
           
        }

        internal void ReportPart(ref iTextSharp.text.Document doc, int partNum)
        {
            // Generate the report part and add it to the document

            SqlDataReader reader = getReportPartData(partNum);
            
           //The title for the report part is a string whose value is in the title collection with the key that corresponds to the enum value for this part number
            //Ex part 1 enum value="There" and title("There")="Part I: What’s There"
          iTextSharp.text.Chunk partTitle=new iTextSharp.text.Chunk((string)title[((myCatetories)partNum-1).ToString()]);
            //the destination is the same as the enum string value for partNum
          partTitle.SetLocalDestination(((myCatetories)partNum-1).ToString());


          iTextSharp.text.Paragraph p = new iTextSharp.text.Paragraph(partTitle);
          p.SpacingAfter = 15;
           doc.Add(p);

            //The description is based on the description collection calculated the same way as the title
           iTextSharp.text.Chunk partdescription = new iTextSharp.text.Chunk((string)description[((myCatetories)partNum-1).ToString()]);
           doc.Add(partdescription);

            //at last - the data table
           iTextSharp.text.pdf.PdfPTable table = new iTextSharp.text.pdf.PdfPTable(reader.FieldCount);
            //put header row on each page
            table.HeaderRows = 1;

        

                for (int i = 0; i <= reader.FieldCount - 1; i++ )
                {
                    //the field names are the headers
                    table.AddCell(reader.GetName(i));
                }
              
            
            if (reader.HasRows)
            {
                while (reader.Read())
                {                

                    for (int i = 0; i <= reader.FieldCount - 1; i++)
                    {
                        table.AddCell(reader[i].ToString());
                    }
                }
            }
            doc.Add(table);
            reader.Close();
           
        }

        internal SqlDataReader getReportPartData(int reportPart)
            {
                try{
               
                    StringBuilder sbCmd = new StringBuilder();
                    string[] partFields;
                    //The list of fields is contained in an array.  The specific array for this part is accessed by the "fields" hashtable using the report part as the key
                    partFields = (string[])fields[reportPart.ToString()];

                    sbCmd.Append("select ");
                    for (int i = 0; i < partFields.Length; i++)
                    {
                        if (i > 0)
                        {
                            sbCmd.Append(",");
                        }
                        sbCmd.Append(partFields[i].Replace(" ", "_") + " as [" + partFields[i]+"]");
                    }
                
                    sbCmd.Append(" from TEMP_TECH_INVENTORY where parent_log_id=@logID and CATEGORY=@category");
                    sbCmd.Append(" order by ");
                    sbCmd.Append(sortby[reportPart - 1]);
                    SqlConnection cnn = dbConnect();
                    SqlCommand thisCommand = new SqlCommand(sbCmd.ToString(), cnn);

                   thisCommand.CommandTimeout = 120;
                   thisCommand.Parameters.Add("@logID", System.Data.SqlDbType.Int);
                   thisCommand.Parameters["@logID"].Value = LogID;
                   thisCommand.Parameters.Add("@category", System.Data.SqlDbType.Int);
                   thisCommand.Parameters["@category"].Value = reportPart;
           
                   SqlDataReader thisReader = thisCommand.ExecuteReader();              
                   return thisReader;

          
                }
                catch (SqlException e)
        	    {
                    Console.WriteLine(e.Message);                
                    return null;

                }
            }

            private static SqlConnection dbConnect()
            {
                try
                {
                    SqlConnection thisConnection = new SqlConnection(@"Data Source=co-pi-sh01-d01,7001;Initial Catalog=Techxl;Integrated Security=True");
                    thisConnection.Open();
                    return thisConnection;

                }
                catch (Exception e )
                {
                    //TODO:log this
                    string m = e.Message;
                    string s = e.Source;
                    Console.WriteLine(s);
                    Console.WriteLine(m);
                    return null;
                }
            }
          
        
        }




    }
