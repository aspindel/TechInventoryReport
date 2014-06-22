using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using iTextSharp;
using iTextSharp.text;
using System.IO;
using System.Data.SqlClient;


namespace TechInventoryReport
{
    class Program
    {        
        string mOracleid;

        internal string OracleID
    {
        get
        {
            return mOracleid;
        }
        set
        {
            mOracleid = OracleID;
        }
    }
 

        static void Main(string[] args)
        {
            string err = "";
            string LogDateTime = DateTime.Now.Year.ToString() + padZero(DateTime.Now.Month.ToString()) + padZero(DateTime.Now.Day.ToString());
            string logfile = LogDateTime + "TechInventoryReport.log";
            SqlConnection cnn = dbConnect(logfile);
            if (cnn == null)
            {
                //Can we connect to db? If not, try to write the error info and quit
                try
                {
                    Logging sqlerrlog = new Logging(logfile);
                    
                    sqlerrlog.LogThis("Unable to connect to database.\r\n" + err + "\r\nQuitting....");
                    
                }
                catch { }
               
                    System.Environment.Exit(1);
                    return;                
            }
            Logging log = new Logging(logfile);

            //Get a data reader of most recently added records in TECH_INVENTORY_HISTORY for each Oracle ID, which is a list of the reports to generate.             

            string OracleID;           
            string deptName="";
            int numThere=0;
            int numMissing = 0;
            int numNever = 0;
            int numElsewhere = 0;
            int numEOLSeen = 0;
            DateTime reportDate;
           
            StringBuilder sbCmd = new StringBuilder();
            sbCmd.Append("select h.*, DEPARTMENT_NICKNAME, DEPARTMENT_NAME from TECH_INVENTORY_HISTORY h ");
            sbCmd.Append("inner join (select max(run_date) as run_date, oracle_id from TECH_INVENTORY_HISTORY group by oracle_id) g ");
            sbCmd.Append("on g.oracle_id=h.oracle_id  and g.run_date=h.RUN_DATE ");
            sbCmd.Append("inner join TECHXL_DEPARTMENTS d on d.ORACLE_ID=h.ORACLE_ID ");
            sbCmd.Append("where REPORT_CREATED=0 and h.run_date=g.run_date ");
#if !DEBUG
            
     sbCmd.Append("and h.run_date>dateadd(hh,-4,getdate()) ");
#endif
            sbCmd.Append("and exists (select * from  TEMP_TECH_INVENTORY t where t.PARENT_LOG_ID=h.TECH_INVENTORY_LOG_ID)");

            log.LogThis("Getting list of reports to generate");
            SqlCommand thisCommand = new SqlCommand(sbCmd.ToString(), cnn);
            int rowsCreated;
            thisCommand.CommandTimeout = 120;
            SqlDataReader thisReader = thisCommand.ExecuteReader();
            int i,LogID,reportCount;
            String filename;
            reportCount = 0;

            if (thisReader.HasRows)
            {
                log.LogThis("Created list of reports to generate.");
                //Generate the reports by iterating through the list.
                while (thisReader.Read())
                {
                    deptName = thisReader["DEPARTMENT_NAME"].ToString();
                    DateTime D;
                    if (DateTime.TryParse(thisReader["RUN_DATE"].ToString(), out D))
                    {
                        reportDate = D;
                    }
                    else
                    {
                        reportDate = DateTime.Now;
                    }
                    if (int.TryParse(thisReader["TECH_INVENTORY_LOG_ID"].ToString(), out i))
                    {

                        //data exist for this Oracle ID
                        LogID = i;
                        numThere = int.Parse(thisReader["SEEN"].ToString());
                        numMissing = int.Parse(thisReader["MISSING"].ToString());
                        numNever = int.Parse(thisReader["NEVER_SEEN"].ToString());
                        numElsewhere = int.Parse(thisReader["SEEN_ELSEWHERE"].ToString());
                        numEOLSeen = int.Parse(thisReader["EOL_SEEN"].ToString());
                        OracleID = thisReader["ORACLE_ID"].ToString();
                        deptName = thisReader["DEPARTMENT_NICKNAME"].ToString();

                        
#if DEBUG
                        filename =  OracleID + "-Tech Inventory-" + LogDateTime + ".pdf";
#else
                        //TODO: make sure z is mapped to https://ssc.cps.edu/Tech_Inventory/ and map it if not
                        filename ="Z:\\" +  OracleID + "-Tech Inventory-" + DateTime.reportDate.Year.ToString() + padZero(DateTime.reportDate.Month.ToString()) + padZero(DateTime.reportDate.Day.ToString() + ".pdf";
#endif
                        log.LogThis("Starting to create report for " + OracleID + " LogID=" + i.ToString());

                        CreateReportForDept(OracleID, deptName, numThere, numMissing, numNever, numElsewhere, numEOLSeen, LogDateTime, LogID);

                        //After generating each report rename the file and set report_created=1 for that ID
                        log.LogThis("Report for " + OracleID + " LogID=" + i.ToString() + " finished.");
                        try
                        {
                            if (System.IO.File.Exists(filename))
                            {
                                System.IO.File.Delete(filename);
                            }
                            System.IO.File.Move("iTextSharpTest.pdf", filename);


                            SqlConnection createcnn = dbConnect(logfile);
                            SqlCommand createdCommand = new SqlCommand("update TECH_INVENTORY_HISTORY set REPORT_CREATED=1 where TECH_INVENTORY_LOG_ID=@LogID", createcnn);

                            createdCommand.Parameters.Add("@LogID", System.Data.SqlDbType.Int);
                            createdCommand.Parameters["@LogID"].Value = i.ToString();

                            int rowsUpdated;
                            rowsUpdated = createdCommand.ExecuteNonQuery();
                            createdCommand.Dispose();
                          
                            createcnn.Dispose();
                            if (rowsUpdated <1)
                            {
                                log.LogThis("WARNING: REPORT_CREATED should be set to 1 for " + i.ToString());
                            }

                            else if (rowsUpdated>1)
                            {
                                log.LogThis("WARNING: there were " + rowsUpdated.ToString() + " rows updated when setting REPORT_CREATED to 1 for record " + i.ToString());
                            } else
                            {
                                log.LogThis("Report renamed & REPORT_CREATED set to 1 for " + i.ToString(),false);
                                log.LogThis("Moving data to TECH_INVENTORY_DATA",false);

                                try
                                {
                                    SqlConnection archivecnn = dbConnect(logfile);
                                    SqlCommand archiveCommand = new SqlCommand("INSERT INTO TECH_INVENTORY_DATA select * from TEMP_TECH_INVENTORY where PARENT_LOG_ID=" + i.ToString(), archivecnn);
                                   

                                    rowsCreated = archiveCommand.ExecuteNonQuery();
                                    if (rowsCreated == numThere + numMissing + numNever + numElsewhere + numEOLSeen)
                                    {
                                        log.LogThis("Data moved to TECH_INVENTORY_DATA", false);
                                        archiveCommand.CommandText = "DELETE from TEMP_TECH_INVENTORY where PARENT_LOG_ID=" + i.ToString();
                                        archiveCommand.ExecuteNonQuery();
                                        reportCount++;
                                    }
                                    else
                                    {
                                        log.LogThis("WARNING: Data NOT moved to TECH_INVENTORY_DATA for " + i.ToString() + " Rowcount<>" + rowsCreated.ToString());                                        
                                        archiveCommand.CommandText = "DELETE from TECH_INVENTORY_DATA where PARENT_LOG_ID=" + i.ToString();
                                        archiveCommand.ExecuteNonQuery();

                                    }
                                   
                                    archiveCommand.Dispose();
                                    
                                    archivecnn.Dispose();
                                }
                                catch (Exception e)
                                {
                                    log.LogThis("WARNING: Data NOT moved to TECH_INVENTORY_DATA for " + i.ToString() + ", " +e.Message );   
                                }

                            }
                        }
                        catch (Exception e)
                        {
                            log.LogThis("WARNING: report created, but not renamed/status not updated for " + i.ToString());
                            log.LogThis(e.Source + "\r\n" + e.Message);
                        }
                    }
                }//end loop through reader                             
            }
            else
            {
                log.LogThis("No reports to generate.");
            }
 
            thisReader.Dispose();
            thisCommand.Dispose();
            cnn.Dispose();
            
          }

        
      

        private static bool CreateReportForDept(string OracleID, string deptName, int numThere, int numMissing, int numNever, int numElsewhere, int numEOLSeen, string LogDateTime, int logID)
        {

            Document doc = new iTextSharp.text.Document();
            FileStream fs = new FileStream("iTextSharpTest.pdf", FileMode.Create, FileAccess.Write, FileShare.None);
            iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, fs);

            //TODO: Add bookmoarks at the start of every section 
            //Open PDF reader with the bookmark panel open 
           // writer.ViewerPreferences = iTextSharp.text.pdf.PdfWriter.PageModeUseOutlines;  
            //NOTES for bookmark adding page height is  715.245 http://stackoverflow.com/questions/19360946/bookmark-to-specific-page-using-itextsharp-4-1-6

            doc.Open();

            addLetter(ref doc);
            doc.NewPage();
            addSummary(deptName,  numThere,  numMissing,  numNever,  numElsewhere,  numEOLSeen,ref doc);
            doc.NewPage();
            addInstructions(ref doc);
            
            reportdata data = new reportdata(OracleID, LogDateTime, doc, logID);
            for  (int i = 1; i <= 5; i++ )
            {
                doc.NewPage();
                data.ReportPart(ref doc, i);
                
            }
            writer.Flush();
            doc.Close();

            return true;

        }

        private static string padZero(string datepart)
        {
            if (datepart.Length == 1)
            {
                datepart = "0" + datepart;
            }
            return datepart;
        }
       

        

        private static void addLetter(ref Document doc){
            //Add the letter(first page) to the report.  This page is generic.
            Font link = FontFactory.GetFont("Arial", 12f, Font.UNDERLINE, BaseColor.BLUE);
            Font normal = FontFactory.GetFont("Arial", 12f);
            Paragraph salutation = new Paragraph("Dear Principals,");
            salutation.SpacingAfter = 10;
            doc.Add(salutation);

            Paragraph letterP1 = new Paragraph("To help further ensure that technology inventory information is accurate and complete and to help reduce computer loss throughout the district, please find enclosed the new technology inventory report for your school. The report contains the following sections:");
            letterP1.SpacingAfter = 10;
            doc.Add(new Paragraph(letterP1));

            iTextSharp.text.List sectionList = new iTextSharp.text.List(iTextSharp.text.List.UNORDERED, 10f);
            sectionList.SetListSymbol("\u2022");           
            sectionList.IndentationLeft = 30f;

            
            Chunk summary = new Chunk("Summary");
            summary.SetLocalGoto("Summary");
            summary.Font = link;
            Phrase sum = new Phrase();
            sum.Add(summary);
            ListItem item = new iTextSharp.text.ListItem();
            item.Font = normal;
            item.Add(sum);
            sectionList.Add(item);

            Chunk There = new Chunk("Part I - What’s There");
            There.SetLocalGoto("There");
            There.Font = link;
            sectionList.Add(new iTextSharp.text.ListItem(There));
            Chunk Missing = new Chunk("Part II - What’s Missing");
            Missing.SetLocalGoto("Missing");
            Missing.Font = link;
            sectionList.Add(new iTextSharp.text.ListItem(Missing));
            Chunk Never = new Chunk("Part III - Never Seen (on CPS Network)");
            Never.SetLocalGoto("Never");
            Never.Font = link;
            sectionList.Add(new iTextSharp.text.ListItem(Never));

            Chunk Elsewhere = new Chunk("Part IV - Seen Elsewhere (on CPS Network)");
            Elsewhere.SetLocalGoto("Elsewhere");
            Elsewhere.Font = link;
            sectionList.Add(new iTextSharp.text.ListItem(Elsewhere));

            Chunk EOL = new Chunk("Part V - Marked as Stolen/Disposed but still seen (on CPS Network)");
            EOL.SetLocalGoto("EOLSeen");
            EOL.Font = link;
            sectionList.Add(new iTextSharp.text.ListItem(EOL));
         
            doc.Add(sectionList);

            Paragraph letterP3 = new Paragraph();
            Chunk P3C1 = new Chunk("The report includes definitions of each part listed above and provides a comprehensive list of the technology assets that are associated with each definition. Also provided are ");
            Chunk P3C2 = new Chunk("detailed instructions");
            Chunk P3C3 = new Chunk(" on how to update your technology inventory and how best to address errors on the technology inventory report.");
            P3C2.Font = link;
            P3C2.SetLocalGoto("instructions");
            letterP3.Add(P3C1);
            letterP3.Add(P3C2);
            letterP3.Add(P3C3);
            letterP3.SpacingBefore = 20;
            letterP3.SpacingAfter = 10;
            doc.Add(letterP3);



            Paragraph letterP4 = new Paragraph("Please review your inventory report and take the necessary steps to reconcile the information. ");
            letterP4.SpacingAfter = 10;           
            doc.Add(letterP4);

            Paragraph letterQuestions = new Paragraph("Answers to frequently asked questions related to the report are on the Knowledge Center at ");
            ///cps.edu / tech - acquisitions.
            Chunk knowledgeCenter = new Chunk("cps.edu/tech-acquisitions");
            knowledgeCenter.SetAnchor("http://cps.edu/tech-acquisitions");
            knowledgeCenter.Font = link;
            letterQuestions.Add(knowledgeCenter);
             letterQuestions.Add(". If you have other questions, please call the School Support Center (SSC) at 773-535-5800.");
            letterQuestions.SpacingAfter = 10;
            doc.Add(letterQuestions);

            Paragraph letterClosing = new Paragraph("Thank you, \nInformation & Technology Services\nin partnership with the School Support Center");
            letterClosing.SpacingAfter = 10;
            doc.Add(letterClosing);
        }


        private static void addInstructions(ref Document doc)
        {
            //Add the instructions to the report.  This page is generic/static.

            Font link = FontFactory.GetFont("Arial", 12f, Font.UNDERLINE, BaseColor.BLUE);
            Chunk title = new Chunk("Detailed Instructions: Part III - Never Seen");
            title.SetLocalDestination("instructions");
            title.Font.SetStyle("bold");
            Paragraph pTitle = new Paragraph();
            pTitle.Alignment = Element.ALIGN_CENTER;
            pTitle.Add(title);            
            doc.Add(pTitle);

            Chunk start = new Chunk("Bring your technology assets under Part III - Never Seen into compliance by completing the following steps:");          
            doc.Add(start);



            Paragraph P2C1 = new Paragraph("Windows machines");
            P2C1.Font.SetStyle("bold");
            P2C1.SpacingAfter = 10;
            doc.Add(P2C1);



            
            iTextSharp.text.List WinInstr = new iTextSharp.text.List(iTextSharp.text.List.UNORDERED, 10f);
            WinInstr.IndentationLeft = 20f;
            WinInstr.SetListSymbol("\u2022");
            Chunk P2C2=new Chunk("Go to ");
            Chunk P2C3=new Chunk("http://school-adm01.instr.cps.k12.il.us");
            P2C3.Font = link;
            P2C3.SetAnchor("http://school-adm01.instr.cps.k12.il.us");    
            Chunk P2C4 =new Chunk(" and run the following tools (in this order):");
            Paragraph p2=new Paragraph();
            //p2.Add(P2C1);
            p2.Add(P2C2);
            p2.Add(P2C3);
            p2.Add(P2C4);
            WinInstr.Add(new iTextSharp.text.ListItem(p2));
            

            iTextSharp.text.List WinInstrList = new iTextSharp.text.List(iTextSharp.text.List.UNORDERED, 10f);
            
            WinInstrList.IndentationLeft = 30f;
            WinInstrList.Add("Correct Computer Name");
            WinInstrList.Add("Asset Utility");  
            WinInstrList.Add("Join Domain tool");
            WinInstrList.Add("Wireless Configuration (if applicable)");
            WinInstrList.Add("Install SCCM Client");
            WinInstrList.Add("Install Anti-virus");
            WinInstrList.Add("Remediation for your operating system (Windows XP, 2000, 2003 or Windows 7)");
            WinInstr.Add(WinInstrList);

            WinInstr.Add(new iTextSharp.text.ListItem("For laptops, please call the ITS Service Desk at (773) 553-3925, option 9, for anti-theft software."));
            doc.Add(WinInstr);
            


            Paragraph P5 = new Paragraph("Mac machines");
            P5.Font.SetStyle("bold");
            P5.SpacingAfter = 10;
            P5.SpacingBefore = 10;
            doc.Add(P5);
            

            Chunk p6c1=new Chunk("Go to ");
            Chunk P6C2=new Chunk("http://school-adm01.instr.cps.k12.il.us");
            P6C2.Font = link;
            P6C2.SetAnchor("http://school-adm01.instr.cps.k12.il.us");    
            Chunk P6C3 =new Chunk(" and run all of the Mandatory Utilities and Maintenance Tools.");
            Paragraph P6=new Paragraph();
            
            P6.Add(p6c1);
            P6.Add(P6C2);
            P6.Add(P6C3);
            iTextSharp.text.List MacList = new iTextSharp.text.List(iTextSharp.text.List.UNORDERED, 10f);
            MacList.SetListSymbol("\u2022");
            MacList.IndentationLeft = 20;
            MacList.Add(new iTextSharp.text.ListItem(P6));



            MacList.Add(new iTextSharp.text.ListItem("For laptops, please call the ITS Service Desk at (773) 553-3925, option 9, for anti-theft software."));

            doc.Add(MacList);


            Paragraph P8 = new Paragraph("iPads");
            P8.Font.SetStyle("bold");
            P8.SpacingAfter = 10;
            P8.SpacingBefore = 10;
            doc.Add(P8);

            iTextSharp.text.List iPad = new iTextSharp.text.List(iTextSharp.text.List.UNORDERED, 10f);
            iPad.SetListSymbol("\u2022");
            iPad.IndentationLeft = 20;
            iPad.Add(new iTextSharp.text.ListItem("Rename the iPad according to the CPS naming convention. "));
          //  doc.Add(iPad);

            iTextSharp.text.List iPadInstr= new iTextSharp.text.List(iTextSharp.text.List.UNORDERED, 10f);

            iPadInstr.IndentationLeft = 30f;
            iPadInstr.Add(new iTextSharp.text.ListItem("Go to the settings page"));
            iPadInstr.Add(new iTextSharp.text.ListItem("Select the General settings"));
            iPadInstr.Add(new iTextSharp.text.ListItem("Select About"));
            iPadInstr.Add(new iTextSharp.text.ListItem("Select the Name field"));
            iPadInstr.Add(new iTextSharp.text.ListItem("Erase the current name and replace it with the new name"));
            iPadInstr.Add(new iTextSharp.text.ListItem("Click Done when finished"));
            iPad.Add(iPadInstr);

           // doc.Add(iPadInstr);

           Paragraph naming=new Paragraph("The first character of the name will be an A or I depending on the primary use of the iPad (instructional or administrative). The next 5 characters of the name are the Oracle financial unit that owns the device. The next character is T for tablet form factor for an iPad. The name ends with the asset tag.");
           naming.IndentationLeft = 30f;
           
           doc.Add(iPad);
           doc.Add(naming);

           Chunk MDM1 = new Chunk("If your department has been set up for AirWatch Mobile Device Management (MDM), enroll the iPads. Instructions for enrollment are available at ");
           Chunk MDM2 = new Chunk("http://ipads.cps.edu");
           MDM2.Font = link;
           MDM2.SetAnchor("http://ipads.cps.edu");
           Chunk MDM3 = new Chunk(". To find out if AirWatch has been activated for your school, please email ipads@cps.edu.  Otherwise, enter their information into FAA manually. Double check the serial number by entering it into ");
           Chunk MDM4 = new Chunk("https://selfsolve.apple.com");
            MDM4.Font=link; 
            MDM4.SetAnchor("https://selfsolve.apple.com");
            Chunk MDM5=new Chunk(" and then paste it into FAA if it is correct. Additional model and warranty information may be available from ");
            Chunk MDM6 = new Chunk(MDM2); //Adding the same chunk twice to the same phrase does not work.
            Chunk MDM7=new Chunk(" for entry into FAA.  For any additional questions about iPads, please go to ");
            Chunk MDM8=new Chunk("http://ipads.cps.edu");
            MDM8.Font = link;
            MDM8.SetAnchor("http://ipads.cps.edu");
            Chunk MDM9=new Chunk(".");

            Phrase MDM = new Phrase(); 
            MDM.Add(MDM1);
            MDM.Add(MDM2);
            MDM.Add(MDM3);
            MDM.Add(MDM4);
            MDM.Add(MDM5);
            MDM.Add(MDM6);
            MDM.Add(MDM7);
            MDM.Add(MDM8);
            MDM.Add(MDM9);
            
            //iPad.Add(new iTextSharp.text.ListItem(MDM));           
            //doc.Add(iPad);
            iTextSharp.text.List MDMListItem = new iTextSharp.text.List(iTextSharp.text.List.UNORDERED, 10f);
            MDMListItem.SetListSymbol("\u2022");
            MDMListItem.IndentationLeft = 20;
            MDMListItem.Add(new iTextSharp.text.ListItem(MDM));
            doc.Add(MDMListItem);
            doc.NewPage();
            Paragraph P9 = new Paragraph("Chromebooks");
            P9.Font.SetStyle("bold");
            P9.SpacingAfter = 10;
            P9.SpacingBefore = 10;
            doc.Add(P9);

            Paragraph P10 = new Paragraph("To enroll your Chromebook under management by cps.edu, please follow the steps below. ");
            P10.SpacingAfter = 10;
            doc.Add(P10);

            Paragraph P11 = new Paragraph();
            Chunk p11_1=new Chunk("If you have ");
            Chunk p11_2=new Chunk("ever");
            p11_2.Font.SetStyle("bold");
            Chunk p11_3=new Chunk(" used the Chromebook, you must reset/restore it before following the steps below to enroll your Chromebook under management by cps.edu. The procedure for resetting/restoring the Chromebook varies by model. ");
            P11.Add(p11_1);
            P11.Add(p11_2);
            P11.Add(p11_3);

            P11.SpacingAfter = 10;
            doc.Add(P11);


            Paragraph P12 = new Paragraph();
            Chunk p12_1=new Chunk("If you have ");
            Chunk p12_2=new Chunk("never");
            p12_2.Font.SetStyle("bold");
            Chunk p12_3=new Chunk(" used the Chromebook, follow these steps:");
            P12.Add(p12_1);
            P12.Add(p12_2);
            P12.Add(p12_3);
            doc.Add(P12);


            iTextSharp.text.List chrome = new iTextSharp.text.List(iTextSharp.text.List.ORDERED, 20f);
            chrome.SetListSymbol("\u2022");
            chrome.IndentationLeft = 20;

            chrome.Add(new iTextSharp.text.ListItem("Connect to network (wifi/ethernet)"));
            chrome.Add(new iTextSharp.text.ListItem("Accept Google terms"));
            chrome.Add(new iTextSharp.text.ListItem("Enroll in cps.edu domain (control-alt-e to access enterprise enrollment screen)"));
            chrome.Add(new iTextSharp.text.ListItem("Log in with your google@cps.edu account (If you do this BEFORE enrolling you will have to restore the device to enroll it)"));


            doc.Add(chrome);

            
            Chunk chromekc1 = new Chunk("Detailed steps are located on the Knowledge Center at ");
            Chunk chromekc2 = new Chunk("cps.edu/tech-acquisitions.");
            chromekc2.SetAnchor("http://cps.edu/tech-acquisitions");
            chromekc2.Font = link;
            Phrase p13 = new Phrase(chromekc1);
            p13.Add(chromekc2);  
            Paragraph P13 = new Paragraph(p13);
            P13.SpacingAfter = 10;
            P13.SpacingBefore = 10; 
            doc.Add(P13);


            Paragraph P14 = new Paragraph("Still not seen on the CPS network?");
            P14.Font.SetStyle("bold");
            P14.SpacingAfter = 10;
            P14.SpacingBefore = 10;
            doc.Add(P14);


            Paragraph P15 = new Paragraph("If you have followed the steps above for your asset type and your asset is still not seen on the CPS network, it may simply not have been on the network long enough to be seen. Please connect the asset to the network and leave it powered on and connected to the network overnight. If the device is still not seen, please call the ITS Service Desk at (773) 553-3925, option 9.");
            P15.SpacingBefore = 10;
            doc.Add(P15);

        }

        private static void addSummary(string deptName, int numThere, int numMissing, int numNever, int numElsewhere, int numEOLSeen, ref Document doc)
        {
            //The method is static, but the page contains report-specific information passed in the parameters
            Chunk summary=new Chunk("Technology Inventory Report for "+deptName);
            summary.SetLocalDestination("Summary");
            Paragraph title = new Paragraph(summary);
            title.Font.SetStyle("bold underline");
            title.SpacingAfter = 10;
            title.Alignment = Element.ALIGN_CENTER;
            doc.Add(title);

            Chunk Sum1 = new Chunk("Summary");
            Sum1.Font.SetStyle("bold");
            doc.Add(Sum1);

            iTextSharp.text.pdf.PdfPTable table = new iTextSharp.text.pdf.PdfPTable(5);
            iTextSharp.text.pdf.PdfPCell cell = new iTextSharp.text.pdf.PdfPCell();
            cell.AddElement(new iTextSharp.text.Chunk("What’s There"));
            table.AddCell(cell);
            table.AddCell("What's Missing");
            table.AddCell("Never Seen (on CPS Network)");
            table.AddCell("Seen Elsewhere (on CPS Network)");
            table.AddCell("Marked as Stolen/Disposed but still seen (on CPS Network)");
            
            table.AddCell(numThere.ToString());
            table.AddCell(numMissing.ToString());
            table.AddCell(numNever.ToString());
            table.AddCell(numElsewhere.ToString());
            table.AddCell(numEOLSeen.ToString());

            doc.Add(table);
      

            Chunk Sec1_Title = new Chunk("Part I - What’s There");
            Sec1_Title.SetLocalGoto("There");
            Paragraph P1_Title = new Paragraph(Sec1_Title);
            P1_Title.Font.SetStyle("bold");       
            
            doc.Add(P1_Title);

            StringBuilder sb = new StringBuilder();
            sb.Append("Part I is a list of technology assets that are “in good standing” in your school – as in, these assets are");
            sb.Append(" in the Fixed Assets Application (FAA) system (“the inventory”) and they have been seen on the CPS network in the last 30 days (during the school year).");

                

            Paragraph P1_Body=new Paragraph(sb.ToString());
            P1_Body.SpacingAfter = 10;
            doc.Add(P1_Body);
            
            sb.Length=0;
            sb.Append("Although no action is required to reconcile this information, it is suggested that you verify the accuracy of the information ");
            sb.Append("and input missing room types, input missing room numbers, and check room numbers for uniformity ");
            sb.Append("(e.g. room 101 should not have multiple entries such as rm 101 and 101).");


            Paragraph P1_2_Body=new Paragraph(sb.ToString());
            P1_2_Body.SpacingAfter = 10;
            doc.Add(P1_2_Body);

            
            Chunk sec2_title = new Chunk("Part II – What’s Missing");
            sec2_title.SetLocalGoto("Missing");
            Paragraph P2_Title = new Paragraph(sec2_title);
            P2_Title.Font.SetStyle("bold");         
            
            doc.Add(P2_Title);

            sb.Length = 0;
            sb.Append("Part II is a list of technology assets that have been seen on the CPS network at some point in time, but have not been seen on the network in the last 30 days (during the school year).");
            
            Paragraph P2_Body=new Paragraph(sb.ToString());                      
            P2_Body.SpacingAfter = 10;
            doc.Add(P2_Body);

            sb.Length = 0;           
            sb.Append("It is important to determine if these technology assets are still in your school:");

            Paragraph P2_2_Body = new Paragraph(sb.ToString());
            doc.Add(P2_2_Body);
           
            iTextSharp.text.List isInSchool = new iTextSharp.text.List(iTextSharp.text.List.UNORDERED, 10f);
            isInSchool.IndentationLeft = 30f;
            isInSchool.SetListSymbol("\u2022");
            Chunk here = new Chunk("If you still have the technology asset, please determine why it is not in use and follow the appropriate steps below.");
            isInSchool.Add(new iTextSharp.text.ListItem(here));
            doc.Add(isInSchool);

            iTextSharp.text.List missingList = new iTextSharp.text.List(iTextSharp.text.List.UNORDERED, 10f);

            missingList.IndentationLeft = 60f;

             Font link = FontFactory.GetFont("Arial", 12f, Font.UNDERLINE, BaseColor.BLUE);


            Chunk dispose1 = new Chunk("If it is too old, please dispose/recycle it. Dispose/recycle steps are located on the Knowledge Center at ");
            Chunk dispose2 = new Chunk("cps.edu/tech-acquisitions.");
            dispose2.SetAnchor("http://cps.edu/tech-acquisitions");
            dispose2.Font = link;
            Phrase dispose = new Phrase(dispose1);
            dispose.Add(dispose2);     
            missingList.Add(new iTextSharp.text.ListItem(dispose));

            missingList.Add(new iTextSharp.text.ListItem("If it needs repair call the IT Service Desk at (773) 553-3925, option 9.  Ensure you have funds set aside for repair."));
            Chunk donate1 = new Chunk("If your school does not need it, donate it to another school. Donation process steps are located on the Knowledge Center at ");
            Chunk donate2 = new Chunk("cps.edu/tech-acquisitions.");
            donate2.Font=link;
            donate2.SetAnchor("http://cps.edu/tech-acquisitions");
            Phrase donate = new Phrase(donate1);
            donate.Add(donate2);
            missingList.Add(new iTextSharp.text.ListItem(donate));

            
            
            missingList.Add(new iTextSharp.text.ListItem("If your school does need it, please turn the technology asset on so that it can be seen on the network. Please make sure to do this every 15 days."));
            doc.Add(missingList);

            iTextSharp.text.List isInSchool2 = new iTextSharp.text.List(iTextSharp.text.List.UNORDERED, 10f);
            isInSchool2.SetListSymbol("\u2022");
            isInSchool2.IndentationLeft = 30f;
            Chunk nothere1 = new Chunk("If you do not have the technology asset, please report it lost/stolen. Lost/stolen report steps are located on the Knowledge Center at ");
            Chunk nothere2 = new Chunk("cps.edu/tech-acquisitions.");
            nothere2.Font = link;
            nothere2.SetAnchor("http://cps.edu/tech-acquisitions");
            Phrase nothere = new Phrase(nothere1);
            nothere.Add(nothere2);
            isInSchool2.Add(new iTextSharp.text.ListItem(nothere));
            doc.Add(isInSchool2);
            
           


            Chunk sec3_title = new Chunk("Part III – Never Seen");
            sec3_title.SetLocalGoto("Never");
            Paragraph P3_Title = new Paragraph(sec3_title);
            P3_Title.Font.SetStyle("bold");
            P3_Title.SpacingBefore = 10;
            doc.Add(P3_Title);

            sb.Length = 0;
            sb.Append("Part III is a list of technology assets that are in your school’s inventory, but that have never been seen on the CPS network. ");
            Paragraph P3_1_Body=new Paragraph(sb.ToString());
            P3_1_Body.SpacingAfter=10;
            doc.Add(P3_1_Body);

            sb.Length = 0;
            sb.Append("It is possible that some of these assets are ");
            sb.Append("mis-categorized and are not actually computers (for example, a projector or kitchen equipment). ");
            sb.Append("Please make sure these technology assets are categorized appropriately.");
            Paragraph P3_2_Body=new Paragraph(sb.ToString());
            P3_2_Body.SpacingAfter=10;
            doc.Add(P3_2_Body);

            sb.Length = 0;
            sb.Append("For the technology assets that are actually computers, there are various reasons they have never been seen on the network. ");
            sb.Append("Most likely they do not have the appropriate software for their operating system. Please follow the ");
            Chunk instr1 = new Chunk(sb.ToString());
            Chunk instr2 = new Chunk("detailed instructions");
            instr2.Font = link;
            instr2.SetLocalGoto("instructions");

            Chunk instr3 = new Chunk(" below so they will be seen in the future.");
            Paragraph P3_3_Body = new Paragraph();
            P3_3_Body.Add(instr1);
            P3_3_Body.Add(instr2);
            P3_3_Body.Add(instr3);
            P3_3_Body.SpacingAfter = 10;
            doc.Add(P3_3_Body);


            Chunk sec4_title = new Chunk("Part IV – Seen Elsewhere");
            sec4_title.SetLocalGoto("Elsewhere");
            Paragraph p4_title = new Paragraph(sec4_title);
            p4_title.Font.SetStyle("bold");
            doc.Add(p4_title);


            Paragraph P4_1_Body = new Paragraph("Part IV is a list of technology assets that are in your school’s inventory, but that have recently been seen elsewhere on the CPS network.  ");
            P4_1_Body.SpacingAfter = 10;
            doc.Add(P4_1_Body);


            Paragraph P4_2_Body = new Paragraph("It is possible that some of these assets require asset transfers to be initiated and/or approved in FAA:");
            doc.Add(P4_2_Body);
            iTextSharp.text.List transfer = new iTextSharp.text.List(iTextSharp.text.List.UNORDERED, 10f);
            transfer.IndentationLeft = 30f;
            transfer.SetListSymbol("\u2022");
            Chunk notInitiated = new Chunk("If an asset transfer is required but your school has not initiated it in FAA, please initiate it in FAA. Please note all new transfers routing for approval now have an auto time out period of 30 days (from receipt of notification).");
            transfer.Add(new iTextSharp.text.ListItem(notInitiated));
            transfer.Add(new iTextSharp.text.ListItem("If your school has initiated an asset transfer in FAA and it is pending your school’s approval, please request approval. "));
            transfer.Add(new iTextSharp.text.ListItem("If your school has approved an asset transfer in FAA and it is pending the receiving school’s receipt, please contact the receiving school and request approval.  "));
            doc.Add(transfer);

            Paragraph P4_3_Body = new Paragraph("If you do not know the contact at the location listed in Part IV, please call the School Support Center (SSC) at 773-535-5800.");
            P4_3_Body.SpacingAfter=10;
            P4_3_Body.SpacingBefore=10;
            doc.Add(P4_3_Body);

            Paragraph P4_4_Body = new Paragraph("It is possible that some of these assets are at your location. If an asset is at your location, please connect it to the network and leave it powered on and connected to the network overnight. The asset should be seen on the CPS network and should be in good standing on your next report.");
            P4_4_Body.SpacingAfter=10;
            doc.Add(P4_4_Body);

            Chunk sec5_title = new Chunk("Part V – Marked as Stolen/Disposed but still seen");
            sec5_title.SetLocalGoto("EOLSeen");
            Paragraph p5_title = new Paragraph(sec5_title);
            p5_title.Font.SetStyle("bold");
            doc.Add(p5_title);

            
            
            Paragraph P5_1_Body = new Paragraph("Part V is a list of technology assets that have been stolen or disposed according to your school’s inventory, but that have been seen on the network since the date stolen/disposed. ");
            P5_1_Body.SpacingAfter = 10;
            doc.Add(P5_1_Body);

            Paragraph P5_2_Body = new Paragraph("It is possible that some of these assets require updates or approvals in FAA. If appropriate, please update “Disposal/Transfer/Inactivate Reason” and “Disposal/Transfer/Inactivate Date” and request approval.");
            P5_2_Body.SpacingAfter = 10;
            doc.Add(P5_2_Body);
            


        }
        
     
        private static SqlConnection dbConnect(string logfile)
        {
            try
            {
                SqlConnection thisConnection = new SqlConnection(@"Data Source=co-pi-sh01-d01,7001;Initial Catalog=Techxl;Integrated Security=True");

                thisConnection.Open();
                return thisConnection;

            }
            catch (Exception e)
            {
                                
                string err = e.Source + "\r\n" + e.Message;
                Console.WriteLine(e.Source);
                Console.WriteLine(e.Message);
                Logging sqlerrlog = new Logging(logfile);

                sqlerrlog.LogThis("Unable to connect to database.\r\n" + err + "\r\nQuitting....");
                return null;
            }
        }
    }
    class Logging
    {
        private string mLogFileName;
        internal String LogFileName
        {
            get
            {
                return mLogFileName;
            }
            
        }
        internal Logging(string fileName)
        {
            mLogFileName = fileName;

        }
        internal void LogThis(string msg, bool full = true)
        {
            using (StreamWriter sw = File.AppendText(LogFileName))
            {
                Log(msg, sw);
            }


        }
        private void Log(string logMessage, TextWriter txtWriter, bool full=true)
        {
            try
            {
                txtWriter.Write("\r\n ");
                if (full)
                {
                    txtWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
                    txtWriter.WriteLine("  :");
                }
                txtWriter.WriteLine("  :{0}", logMessage);
                if (full)
                {
                    txtWriter.WriteLine("-----------\r\n");
                }
            }
            catch (Exception ex)
            {
                //if you can't log the error, what do you do?
                Console.WriteLine(ex.InnerException);
            }

        }
    }
}
