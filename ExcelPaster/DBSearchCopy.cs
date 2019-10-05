using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelPaster
{
    public class DBSearchCopy
    {
        public void StartSearchCopy(string source, string target,string outputLoc)
        {
            CSVReader sourceReader = new CSVReader();
            CSVReader targetReader = new CSVReader();

            sourceReader.ParseCSV(source,"}");
            targetReader.ParseCSV(target,"");

            List<List<string>> sourceList = new List<List<string>>();
            List<List<string>> targetList = new List<List<string>>();

            sourceList = sourceReader.GetArrayStorage();
            targetList = targetReader.GetArrayStorage();

            //Tidy up Lists
            List<OrderEntry> orderList = new List<OrderEntry>();
            List<FinalEntry> finalList = new List<FinalEntry>();
            List<OrderEntry> orderFailToMatchjList = new List<OrderEntry>();

            int indexID = 0;
            foreach (List<string> entry in sourceList)
            {
                if (indexID == 0) {
                    indexID = 1;
                    continue;
                }
                indexID++;
               
                OrderEntry oEntry = new OrderEntry();
                oEntry.indexID = indexID;
                oEntry.orderID = Int32.Parse(entry[0]);
                oEntry.custID = Int32.Parse(entry[1]);
                oEntry.shipID = Int32.Parse(entry[3]);
                oEntry.notes = entry[4];
                oEntry.SearchNotes();

                orderList.Add(oEntry);
            }

            indexID = 0;
            foreach (List<string> entry in targetList)
            {
                if (indexID == 0)
                {
                    indexID = 1;
                    continue;
                }
                indexID++;
                //TODO:fix these
                FinalEntry oEntry = new FinalEntry();
                oEntry.indexID = indexID;
                oEntry.APIID = Int64.Parse(entry[0]);
                oEntry.customerName = entry[4];
                oEntry.siteName = entry[5];

                finalList.Add(oEntry);
            }

            //Start finding matches
            int[] matchScore = new int[finalList.Count];
            int successfulMatchCounter = 0;
            int processedAmount = 0;
            foreach (OrderEntry oE in orderList)
            {

                foreach (string name in oE.sitenames)
                {
                    if (!name.Contains("CORIOLIS"))//theres a random scope inside a Service At: field
                    {
                        if (Properties.Settings.Default.UseLevensteins)
                        {
                            //assign scores per name in comments
                            int matchCounter = 0;
                            foreach (FinalEntry fE in finalList)
                            {
                                matchScore[matchCounter] = ComputeDistance(name, fE.siteName);
                                matchCounter++;
                            }
                            //Find lowest scores and assign them as probable matches
                            int bestScore = matchScore.Min();
                            int counter = 0;
                            foreach (int value in matchScore)
                            {
                                if (value == bestScore)
                                {
                                    finalList[counter].proposedSiteIDs.Add(oE.shipID);
                                    finalList[counter].proposedeNameMatch.Add(name);
                                    finalList[counter].proposedWorkOrder.Add(oE.orderID);
                                }
                                counter++;
                            }
                        }
                        else
                        {
                            int counter = 0;
                            foreach (FinalEntry fE in finalList)
                            {
                                bool match = CanContain(name, fE.siteName);
                                if (match)
                                {
                                    finalList[counter].proposedSiteIDs.Add(oE.shipID);
                                    finalList[counter].proposedeNameMatch.Add(name);
                                    finalList[counter].proposedWorkOrder.Add(oE.orderID);

                                    orderList[oE.indexID].matched = true;
                                    successfulMatchCounter++;
                                }
                                
                                counter++;
                            }
                        }
                    }

                }
                processedAmount++;
                
            }
            //Write Results
            string filePath = outputLoc + "/Results.csv";
            string delimiter = ",";

            using (FileStream fs = new FileStream(filePath, FileMode.CreateNew, FileAccess.Write))
            {
                if (fs.CanWrite)
                {
                    string csvInput = "";
                    csvInput = "Index_ID" + delimiter + "API_ID" + delimiter + "Customer" + delimiter + "Site_Names" + delimiter + "Prop_Ship_ID1" + 
                        delimiter + "Prop_Name1" + delimiter + "Prop_Work_Order1" + delimiter + "Match Rate: " + successfulMatchCounter + "/" + processedAmount;
                    fs.Write(Encoding.ASCII.GetBytes(csvInput), 0, ASCIIEncoding.ASCII.GetByteCount(csvInput));
                    byte[] newline = Encoding.ASCII.GetBytes(Environment.NewLine);
                    fs.Write(newline, 0, newline.Length);
                    foreach (FinalEntry fE in finalList)
                    {
                        
                        csvInput = fE.indexID + delimiter + fE.APIID + delimiter + fE.customerName + delimiter + fE.siteName;
                        fs.Write(Encoding.ASCII.GetBytes(csvInput), 0, ASCIIEncoding.ASCII.GetByteCount(csvInput));
                       
                        int count = 0;
                        foreach (int siteID in fE.proposedSiteIDs)
                        {
                            csvInput = delimiter + fE.proposedSiteIDs[count] + delimiter + fE.proposedeNameMatch[count] + delimiter + fE.proposedWorkOrder[count];
                            fs.Write(Encoding.ASCII.GetBytes(csvInput), 0, ASCIIEncoding.ASCII.GetByteCount(csvInput));
                            count++;
                        }
                        
                        fs.Write(newline, 0, newline.Length);
                    }

                   
                }
                fs.Flush();
                fs.Close();
            }
            filePath = outputLoc + "/FailedMatches.csv";
            

            using (FileStream fs = new FileStream(filePath, FileMode.CreateNew, FileAccess.Write))
            {
                if (fs.CanWrite)
                {
                    string csvInput = "";
                    csvInput = "Index_ID" + delimiter + "Order_ID" + delimiter + "Ship_ID" + delimiter + "Site_Names" + delimiter + "Notes";
                    fs.Write(Encoding.ASCII.GetBytes(csvInput), 0, ASCIIEncoding.ASCII.GetByteCount(csvInput));
                    byte[] newline = Encoding.ASCII.GetBytes(Environment.NewLine);
                    fs.Write(newline, 0, newline.Length);

                    foreach (OrderEntry oE in orderList)
                    {
                        if (oE.matched == false)
                        {
                            csvInput = oE.indexID + delimiter + oE.orderID + delimiter + oE.shipID + delimiter + oE.sitenames[0] + delimiter + oE.notes;
                            fs.Write(Encoding.ASCII.GetBytes(csvInput), 0, ASCIIEncoding.ASCII.GetByteCount(csvInput));
                            fs.Write(newline, 0, newline.Length);
                        }
                        
                    }


                }
                fs.Flush();
                fs.Close();
            }





        }
        public class OrderEntry
        {
            public int indexID;
            public int orderID;
            public int custID;
            public int shipID;
            public string notes;
            public void SearchNotes()
            {

                // String[] lines = notes.Split(new String[] { Environment.NewLine }, StringSplitOptions.None);
                String[] split1 = Regex.Split(notes.ToUpper(), @"(?<=SERVICE AT)");
                if (split1[1].Contains(':'))
                {
                    
                    split1[1] = split1[1].Remove(split1[1].IndexOf(':'),1);
                }
                string fullName = "";

                //Cut off end
                bool foundCutOff = false;
                String[] cutOffWords = new String[] { "SERVICE DATE","DATE COMPLETED","WO#","SCOPE","SERVICE BY", "SERVICE COMPLETED" };

                foreach (string cutoffWord in cutOffWords)
                {
                    String[] split2 = split1[1].ToUpper().Split(new[] { cutoffWord }, StringSplitOptions.None);
                    if (split2.Count() > 1)
                    {
                        fullName = split1[0] + ":" + split2[0];
                        break;
                    }
                    
                }

                
                 
                String[] lines = new String[] { fullName };
                int assigningID = 0;
                foreach (string line in lines)
                {
                    if (line.Contains(':'))
                    {
                        String[] noteSplit = line.Split(':');
                        switch (noteSplit[0].ToUpper())
                        {
                            case "SERVICE AT":
                                sitenames = new List<string>();
                                assigningID = 0;
                                break;
                            case "SERVICE DATE":
                                assigningID = 1;
                                break;
                            case "CC/WO":
                                assigningID = 2;
                                break;
                            case "COMPLETE BY":
                                assigningID = 3;
                                break;
                            case "SCOPE":
                                assigningID = 4;
                                break;
                            default:
                                break;
                        }
                        AssignFromNoteID(assigningID, noteSplit[1]);
                        continue;
                    }
                    AssignFromNoteID(assigningID, line);
                }
            }
            private void AssignFromNoteID(int id,string text)
            {
                switch (id)
                {
                    case 0:
                        if (sitenames == null)
                        {
                            sitenames = new List<string>();
                        }
                        sitenames.Add(text);
                        break;
                    case 1:
                        serviveDates = text;
                        break;
                    case 2:
                        CCorWO = Int32.Parse(text);
                        break;
                    case 3:
                        completedBy = text;
                        break;
                    case 4:
                        scope = text;
                        break;
                    default:
                        break;



                }
            }
            public List<String> sitenames;
            public string serviveDates;
            public int CCorWO;
            public string completedBy;
            public string scope;
            public bool matched = false;


        }
        public class FinalEntry
        {
            public int indexID;
            public long APIID;
            public string customerName;
            public string siteName;

            public List<int> proposedSiteIDs = new List<int>();
            public List<string> proposedeNameMatch = new List<string>();
            public List<int> proposedWorkOrder = new List<int>();

        }
        /// <summary>
        /// Compute the distance between two strings.
        /// </summary>
        public int ComputeDistance(string s, string t)
        {
            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            // Step 1
            if (n == 0)
            {
                return m;
            }

            if (m == 0)
            {
                return n;
            }

            // Step 2
            for (int i = 0; i <= n; d[i, 0] = i++)
            {
            }

            for (int j = 0; j <= m; d[0, j] = j++)
            {
            }

            // Step 3
            for (int i = 1; i <= n; i++)
            {
                //Step 4
                for (int j = 1; j <= m; j++)
                {
                    // Step 5
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;

                    // Step 6
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);
                }
            }
            // Step 7
            return d[n, m];
        }

        public bool CanContain(string s, string t)
        {
           

            String[] elements = s.Split(' ');

            //Combine Texts if in order
            List<String> refElements = new List<string>();
            string eleToStore = "";
            foreach (string element in elements)
            {
                if (!element.Any(char.IsDigit))
                {
                    if (eleToStore == "")
                    {
                        eleToStore += element;
                    }
                    else
                    {
                        eleToStore += " " + element;
                       
                    }
                    
                }
                else
                {
                    refElements.Add(eleToStore);
                    eleToStore = "";
                    refElements.Add(element);
                }
            }
                

            double score = 0;
            double totalScore = 0;
            foreach (string element in refElements)
            {
                if (element != "" )
                {
                    if (t.Contains(element))
                    {
                        score += 0.25;
                        if ((" " + t + " ").Contains(" " + element + " "))
                        {
                            score += 0.75;
                        }
                    }
                    totalScore++;
                }
               
            }
            if (score / totalScore > 0.75)//75% accurate
            {

                return true;
            }
            else //Try Bruin naming
            {
                String[] splitByFB = s.Split(new string[] { "FB"}, StringSplitOptions.None);
                if (splitByFB.Count() > 1 )
                {
                    if (splitByFB[1] != "")
                    {
                        for (int i = 1; i < splitByFB.Count(); i++)
                        {
                            string FBName = splitByFB[i];
                            if (FBName[0] == '-')
                            {
                                FBName = ' ' + FBName.Remove(0, 1);
                            }
                            FBName = "FORT BERTHOLD" + FBName;
                            FBName = FBName.Replace(".", "");

                            //Repeat Search of elements
                            String[] elementsBruin = FBName.Split(' ');

                            //Combine Texts if in order
                            List<String> refElementsBruin = new List<string>();
                            string eleToStoreBruin = "";
                            foreach (string element in elementsBruin)
                            {
                                if (!element.Any(char.IsDigit))
                                {
                                    if (eleToStoreBruin == "")
                                    {
                                        eleToStoreBruin += element;
                                    }
                                    else
                                    {
                                        eleToStoreBruin += " " + element;

                                    }

                                }
                                else
                                {
                                    refElementsBruin.Add(eleToStoreBruin);
                                    eleToStoreBruin = "";
                                    refElementsBruin.Add(element);
                                }
                            }


                            double scoreBruin = 0;
                            double totalScoreBruin = 0;
                            foreach (string element in refElementsBruin)
                            {
                                if (element != "")
                                {
                                    if (t.Contains(element))
                                    {
                                        scoreBruin += 0.25;
                                        if ((" " + t + " ").Contains(" " + element + " "))
                                        {
                                            scoreBruin += 0.75;
                                        }
                                    }
                                    totalScoreBruin++;
                                }

                            }
                            if (scoreBruin / totalScoreBruin > 0.75)//75% accurate
                            {

                                return true;
                            }

                        }
                    }
                   
                }
                else
                {
                    return false;
                }
            }
            return false;
        }


    }
}
