﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
namespace TestConsole
{

    class Program
    {
        private const string BA = "Procurement General[de]Allgemeiner Einkauf";
        private const string BC = "Connectivity, eMobility&Fahrerassistenz[de]Connectivitaet, eMobilitaet&Fahrerassistenz";
        private const string BI = "Interior[de]Interieur[es]Interior[pl]wnętrze[ru]интерьер[zh]室内[fr]intérieur[it]interno[pt]interior";
        private const string BE = "Electric[de]Elektrik[es]Eléctrico[pl]Elektryczny[ru]электрический[zh]电动[fr]électrique[it]elettrico[pt]elétrico";
        private const string BM = "Metal[de]Metall[es]metal[pl]metal[ru]металл[zh]金属[fr]métal[it]metallo[pt]metal";
        private const string BN = "New Product Launches[de]Neue Produktanlaeufe";
        private const string BP = "Powertrain[de]Antriebswelle[es]Tren motriz[pl]Powertrain[ru]Силовой агрегат[zh]动力总成[fr]Groupe motopropulseur[it]Powertrain[pt]Powertrain";
        private const string BX = "Exterior[de]Exterieur[es]exterior[pl]powierzchowność[ru]экстерьер[zh]外观[fr]extérieur[it]esterno[pt]exterior";

        private const string A = "Procurement General[de]Allgemeiner Einkauf";
        private const string C = "Connectivity, eMobility&Fahrerassistenz[de]Connectivity, eMobility&Fahrerassistenz";
        private const string E = "Electric[de]Elektrik";
        private const string I = "Interior[de]Interieur";
        private const string L = "HV Batterie[de]HV Battery";
        private const string M = "Metal[de]Metall";
        private const string N = "New Product Launches[de]Neue Produktanlaeufe";
        private const string P = "Powertrain[de]Powertrain";
        private const string S = "Services[de]Services";
        private const string X = "Exterior[de]Exterieur";

        private static class SineqaEventType
        {
            public const string SearchText = "search.text";
            public const string SearchResultLink = "click.resultlink";
            public const string DocPreview = "doc.preview";
        }
        static void Main(string[] args)
        {
            //List<string> test = new List<string>();
            //string t1 = "dupa";
            //if (!test.Contains(t1))
            //{
            //    test.Add(t1);
            //}
            //if (!test.Contains(t1))
            //{
            //    test.Add(t1);
            //}
            //t1 = "blada";
            //if (!test.Contains(t1))
            //{
            //    test.Add(t1);
            //}
            //string t2 = string.Join(";", test.ToArray());

            ReadCsvCommodities2();
            //ReadCsvNewLdbOrgTree();
            //CheckProblems();
            //FixOrphans();
        }

        private static void FixOrphans()
        {
            List<LdbItem> newLDBItems = new List<LdbItem>();

            using (StreamReader sr = new StreamReader(@"C:\kpl\newLdbFull.csv"))
            {

                string currentLine;
                // currentLine will be null when the StreamReader reaches the end of file
                while ((currentLine = sr.ReadLine()) != null)
                {
                    string[] coulumns = currentLine.Split(new char[] { ';' });
                    LdbItem item = new LdbItem();
                    item.Duns = CleanInput(coulumns[0]);
                    item.Name = CleanInput(coulumns[1]);
                    item.HqDuns = CleanInput(coulumns[2]);
                    item.HqName = CleanInput(coulumns[3]);
                    item.GuDuns = CleanInput(coulumns[4]);
                    item.GuName = CleanInput(coulumns[5]);
                    newLDBItems.Add(item);
                }
            }

            List<OrgTreeItem> rootItems = new List<OrgTreeItem>();
            //test orphan
            LdbItem orphanitem = newLDBItems.FirstOrDefault(x => x.Duns.Equals("31-606-7164"));
            OrgTreeItem orgItem = new OrgTreeItem();
            orgItem.duns = orphanitem.Duns;
            orgItem.label = orphanitem.Name;
            rootItems.Add(orgItem);
            List<LdbItem> famillyItems = newLDBItems.Where(x => x.GuDuns.Equals(orgItem.duns)).ToList();
            List<string> famillyDuns = famillyItems.Select(x => x.Duns).ToList();
            List<LdbItem> orphanItems = famillyItems.Where(x => !famillyDuns.Contains(x.HqDuns)).ToList();
            foreach (LdbItem orphan in orphanItems.Where(x => !x.OutOfBusiness).ToList())
            {
                orphan.HqDuns = orphan.GuDuns;
            }
            BuildOrgTree(orgItem, famillyItems, orgItem);

            // get GUCs

            //List<LdbItem> gucItems = newLDBItems.Where(x => x.GuDuns.Equals(x.Duns)).ToList();
            //foreach (LdbItem item in gucItems)
            //{
            //    OrgTreeItem orgItem = new OrgTreeItem();
            //    orgItem.duns = item.Duns;
            //    orgItem.label = item.Name;
            //    rootItems.Add(orgItem);
            //    BuildOrgTree(orgItem, newLDBItems.Where(x => x.GuDuns.Equals(orgItem.duns)).ToList(), orgItem);
            //}
            foreach (OrgTreeItem item in rootItems.Where(x => x.items.Count > 0))
            {
                DisplayWrongItems(item, newLDBItems, item.duns); ;
                //foreach (OrgTreeItem child in item.items)
                //{
                //    LdbItem childItem = newLDBItems.FirstOrDefault(x=> x.duns.Equals(child.duns));
                //    if (childItem != null && !childItem.GuDuns.Equals(item.duns))
                //    {
                //        string wrongItem = $"duns:{childItem.duns}, parent:{childItem.HqDuns}, gud:{childItem.GuDuns}";
                //        Console.WriteLine(wrongItem);
                //    }

                //}
            }
        }

        private static void CheckProblems()
        {
            List<LdbItem> newLDBItems = new List<LdbItem>();

            using (StreamReader sr = new StreamReader(@"C:\kpl\newLdbFull.csv"))
            {

                string currentLine;
                // currentLine will be null when the StreamReader reaches the end of file
                while ((currentLine = sr.ReadLine()) != null)
                {
                    string[] coulumns = currentLine.Split(new char[] { ';' });
                    LdbItem item = new LdbItem();
                    item.Duns = CleanInput(coulumns[0]);
                    item.Name = CleanInput(coulumns[1]);
                    item.HqDuns = CleanInput(coulumns[2]);
                    item.HqName = CleanInput(coulumns[3]);
                    item.GuDuns = CleanInput(coulumns[4]);
                    item.GuName = CleanInput(coulumns[5]);
                    newLDBItems.Add(item);
                }
            }

            // get GUCs

            List<LdbItem> gucItems = newLDBItems.Where(x => x.GuDuns.Equals(x.Duns)).ToList();
            List<OrgTreeItem> rootItems = new List<OrgTreeItem>();
            foreach (LdbItem item in gucItems)
            {
                OrgTreeItem orgItem = new OrgTreeItem();
                orgItem.duns = item.Duns;
                orgItem.label = item.Name;
                rootItems.Add(orgItem);
                BuildOrgTree(orgItem, newLDBItems.Where(x => x.GuDuns.Equals(orgItem.duns)).ToList(), orgItem);
            }
            foreach (OrgTreeItem item in rootItems.Where(x => x.items.Count > 0))
            {
                DisplayWrongItems(item, newLDBItems, item.duns); ;
                //foreach (OrgTreeItem child in item.items)
                //{
                //    LdbItem childItem = newLDBItems.FirstOrDefault(x=> x.duns.Equals(child.duns));
                //    if (childItem != null && !childItem.GuDuns.Equals(item.duns))
                //    {
                //        string wrongItem = $"duns:{childItem.duns}, parent:{childItem.HqDuns}, gud:{childItem.GuDuns}";
                //        Console.WriteLine(wrongItem);
                //    }

                //}
            }
        }

        private static void DisplayWrongItems(OrgTreeItem item, List<LdbItem> newLDBItems, string gud)
        {
            foreach (OrgTreeItem child in item.items)
            {
                LdbItem childItem = newLDBItems.FirstOrDefault(x => x.Duns.Equals(child.duns));
                if (childItem != null && !childItem.GuDuns.Equals(gud))
                {
                    string wrongItem = $"root:{gud} duns:{childItem.Duns}, parent:{childItem.HqDuns}, gud:{childItem.GuDuns}";
                    Console.WriteLine(wrongItem);
                }

                DisplayWrongItems(child, newLDBItems, gud);
            }
        }

        private static void ReadCsvNewLdbOrgTree()
        {
            List<LdbItem> newLDBItems = new List<LdbItem>();

            //using (StreamReader sr = new StreamReader(@"C:\kpl\export3rootCompanyB.csv"))
            //using (StreamReader sr = new StreamReader(@"D:\kpl\export3rootCompanyB.csv"))
            using (StreamReader sr = new StreamReader(@"D:\kpl\exportFriedeSpringer.csv"))
            {

                string currentLine;
                // currentLine will be null when the StreamReader reaches the end of file
                while ((currentLine = sr.ReadLine()) != null)
                {
                    string[] coulumns = currentLine.Split(new char[] { ';' });
                    LdbItem item = new LdbItem();
                    item.Duns = CleanInput(coulumns[0]);
                    item.Name = CleanInput(coulumns[1]);
                    item.HqDuns = CleanInput(coulumns[2]);
                    item.HqName = CleanInput(coulumns[3]);
                    item.GuDuns = CleanInput(coulumns[4]);
                    item.GuName = CleanInput(coulumns[5]);
                    newLDBItems.Add(item);
                }
            }

            // get GUCs

            List<LdbItem> gucItems = newLDBItems.Where(x => x.GuDuns.Equals(x.Duns)).ToList();
            List<OrgTreeItem> rootItems = new List<OrgTreeItem>();
            foreach (LdbItem item in gucItems)
            {
                OrgTreeItem orgItem = new OrgTreeItem();
                orgItem.duns = item.Duns;
                orgItem.label = item.Name;
                rootItems.Add(orgItem);
                List<LdbItem> famillyItems = newLDBItems.Where(x => x.GuDuns.Equals(orgItem.duns)).ToList();
                BuildOrgTree(orgItem, famillyItems, orgItem);
                List<OrgTreeItem> path = new List<OrgTreeItem>();
                path.Add(orgItem);
                PersistOrgTree(orgItem, orgItem, path);
            }
            foreach (OrgTreeItem item in rootItems)
            {
                //PersistOrgTree(item, item);
                //item.selected = true;
                //item.expanded = true;
                //string orgtreeJson = new JavaScriptSerializer().Serialize(item);
                // update item where gud = item.duns and set varchar
            }
        }

        private static void PersistOrgTree(OrgTreeItem item, OrgTreeItem rootItem, List<OrgTreeItem> pathItems)
        {
            item.selected = true;
            item.expanded = true;
            
            string orgtreeJson = new JavaScriptSerializer().Serialize(rootItem);
            // update item where sourcestr70 = item.duns and set varchar
            item.selected = false;
            item.expanded = item.items.Count > 0;
            //if (null != parentItem && item.items.Count == 0)
            if (item.items.Count == 0)
            {
                for (int i = pathItems.Count - 1; i > 0; i--)
                {
                    OrgTreeItem checkItem = pathItems[i];
                    OrgTreeItem parentItem = pathItems[i-1];
                    if (parentItem.items.Last().duns.Equals(checkItem.duns))
                    {
                        parentItem.expanded = false;
                        pathItems.RemoveAt(i);

                    }
                    else
                    {
                        pathItems.RemoveAt(i);
                        break;
                    }
                }

                //pathItems.Clear();
            }

            for (int i=0; i< item.items.Count; i++)
            {
                Console.WriteLine(i);
                pathItems.Add(item.items[i]);
                PersistOrgTree(item.items[i], rootItem, pathItems);

                //if (item.items.Count == i)
                //{
                //    PersistOrgTree(item.items[i], rootItem, pathItems);

                //}
                //else
                //{
                //    PersistOrgTree(item.items[i], rootItem, pathItems);

                //}
            }
        }

        //private static OrgTreeItem FindItemByDuns2(OrgTreeItem item, string duns)
        //{
        //    OrgTreeItem currentItem = item.items.FirstOrDefault(x => x.duns.Equals(duns));
        //    if (currentItem != null)
        //    {
        //        item.expanded = true;
        //        return currentItem;
        //    }
            
        //    foreach (OrgTreeItem child in item.items)
        //    {
        //        return FindItemByDuns(child, duns);
        //    }

        //    return null;
        //}

        private static void FindItemByDuns3(OrgTreeItem item, string duns, OrgTreeItem result)
        {
            OrgTreeItem currentItem = item.items.FirstOrDefault(x => x.duns.Equals(duns));
            if (currentItem != null)
            {
                item.expanded = true;
                result = currentItem;
            }
            else {
                foreach (OrgTreeItem child in item.items)
                {
                    FindItemByDuns(child, duns, result);
                }
            }

        }

        private static OrgTreeItem FindItemByDuns(OrgTreeItem item, string duns, OrgTreeItem result)
        {
            //Console.WriteLine(item.duns);
            if (item.duns.Equals(duns))
            {
                return item;
            }

            foreach (OrgTreeItem child in item.items)
            {
                

                return FindItemByDuns(child, duns, result);
            }

            return null;
        }

        //private void ExpandPath(OrgTreeItem item)
        //{
        //    item.expanded = true;
        //    if (!item.duns.Equals(item.parentDuns))
        //    {
        //        return currentItem;
        //    }

        //    foreach (OrgTreeItem child in item.items)
        //    {
        //        return FindItemByDuns(child, duns);
        //    }

        //    return null;
        //}

        //private static void PersistOrgTree(OrgTreeItem item, OrgTreeItem rootItem)
        //{
        //    OrgTreeItem itemClone = (OrgTreeItem)rootItem.Clone();
        //    item.selected = true;
        //    item.expanded = true;

        //    string orgtreeJson = new JavaScriptSerializer().Serialize(rootItem);
        //    // update item where sourcestr70 = item.duns and set varchar
        //    foreach (OrgTreeItem child in item.items)
        //    {
        //        PersistOrgTree(child, itemClone);
        //    }

        //}
        private static void BuildOrgTree(OrgTreeItem item, List<LdbItem> newLDBItems, OrgTreeItem root)
        {
            List<LdbItem> childItems = newLDBItems.Where(x => x.HqDuns.Equals(item.duns) && !x.Duns.Equals(x.GuDuns)).ToList();
            foreach (LdbItem child in childItems)
            {
                OrgTreeItem childOrgItem = new OrgTreeItem();
                childOrgItem.duns = child.Duns;
                childOrgItem.label = child.Name;
                item.items.Add(childOrgItem);
                root.count++;
                BuildOrgTree(childOrgItem, newLDBItems, root);

            }

        }

        private static void ReadCsvCommodities()
        {
            using (StreamReader sr = new StreamReader(@"C:\kpl\commodities.csv"))
            {
                Regex regex = new Regex(@"\d{4}");
                string currentLine;
                Dictionary<string, Comm> values = new Dictionary<string, Comm>();
                Dictionary<string, string> commodities = new Dictionary<string, string>();
                commodities.Add("BA", BA);
                commodities.Add("BC", BC);
                commodities.Add("BI", BI);
                commodities.Add("BE", BE);
                commodities.Add("BM", BM);
                commodities.Add("BN", BN);
                commodities.Add("BP", BP);
                commodities.Add("BX", BX);
                // currentLine will be null when the StreamReader reaches the end of file
                while ((currentLine = sr.ReadLine()) != null)
                {
                    string[] coulumns = currentLine.Split(new char[] { ';' });
                    Match match = regex.Match(coulumns[1]);
                    if (match.Success)
                    {
                        string commValue = coulumns[0];
                        if (!commodities.ContainsKey(commValue))
                        {
                            continue;
                        }


                        string id = match.Value;
                        Comm comm;
                        if (values.ContainsKey(id))
                        {
                            comm = values[id];
                        }
                        else
                        {
                            comm = new Comm();
                            values.Add(id, comm);
                        }

                        comm.Id = id;
                        comm.Translation = commodities[commValue];
                        //if (coulumns[2].Equals("US", StringComparison.InvariantCultureIgnoreCase))
                        //{
                        //    comm.Us = coulumns[3];
                        //}
                        //else if (coulumns[2].Equals("D", StringComparison.InvariantCultureIgnoreCase))
                        //{
                        //    comm.De = coulumns[3];
                        //}
                    }
                }

                StringBuilder sb = new StringBuilder();
                foreach (Comm item in values.Values)
                {
                    //if (string.IsNullOrEmpty(item.Us))
                    //{
                    //    continue;
                    //}

                    string commodityLine = string.Empty;
                    //if (string.IsNullOrEmpty(item.De)) {
                    // commodityLine = string.Format("{0};{1}", item.Us, item.Id);
                    //}
                    //else
                    //{
                    //    commodityLine = string.Format("{0}[de]{1};{2}", item.Us, item.De, item.Id);
                    //}
                    commodityLine = string.Format("{0};{1}", item.Translation, item.Id);

                    sb.AppendLine(commodityLine);
                }

                File.WriteAllText(@"C:\kpl\commoditiesFile.txt", sb.ToString());
            }
        }


        private static void ReadCsvCommodities2()
        {
            using (StreamReader sr = new StreamReader(@"TestData\commodities20200427.csv"))
            {
                Regex regex = new Regex(@"\d{4}");
                string currentLine;
                Dictionary<string, Comm> values = new Dictionary<string, Comm>();
                Dictionary<string, string> commodities = new Dictionary<string, string>();
                commodities.Add("A", A);
                commodities.Add("C", C);
                commodities.Add("E", E);
                commodities.Add("I", I);
                commodities.Add("L", L);
                commodities.Add("M", M);
                commodities.Add("N", N);
                commodities.Add("P", P);
                commodities.Add("S", S);
                commodities.Add("X", X);
                // currentLine will be null when the StreamReader reaches the end of file
                while ((currentLine = sr.ReadLine()) != null)
                {
                    string[] coulumns = currentLine.Split(new char[] { ';' });
                    Match match = regex.Match(coulumns[0]);
                    if (match.Success)
                    {
                        string commValue = coulumns[1];
                        if (!commodities.ContainsKey(commValue))
                        {
                            continue;
                        }


                        string id = match.Value;
                        Comm comm;
                        if (values.ContainsKey(id))
                        {
                            comm = values[id];
                        }
                        else
                        {
                            comm = new Comm();
                            values.Add(id, comm);
                        }

                        comm.Id = id;
                        comm.Translation = commodities[commValue];
                       
                    }
                }

                StringBuilder sb = new StringBuilder();
                foreach (Comm item in values.Values)
                {
                    //if (string.IsNullOrEmpty(item.Us))
                    //{
                    //    continue;
                    //}

                    string commodityLine = string.Empty;
                    //if (string.IsNullOrEmpty(item.De)) {
                    // commodityLine = string.Format("{0};{1}", item.Us, item.Id);
                    //}
                    //else
                    //{
                    //    commodityLine = string.Format("{0}[de]{1};{2}", item.Us, item.De, item.Id);
                    //}
                    commodityLine = string.Format("{0};{1}", item.Translation, item.Id);

                    sb.AppendLine(commodityLine);
                }

                File.WriteAllText(@"TestData\commoditiesFile.txt", sb.ToString());
            }
        }



        private static void CheckCsvCommodities()
        {
            Regex regex = new Regex(@"\d{4}");
            Dictionary<string, Comm> values = new Dictionary<string, Comm>();
            Dictionary<string, string> commodities = new Dictionary<string, string>();
            commodities.Add("BA", BA);
            commodities.Add("BC", BC);
            commodities.Add("BI", BI);
            commodities.Add("BE", BE);
            commodities.Add("BM", BM);
            commodities.Add("BN", BN);
            commodities.Add("BP", BP);
            commodities.Add("BX", BX);
            StringBuilder sb = new StringBuilder();
            List<string> matGroup = new List<string>();

            using (StreamReader sr = new StreamReader(@"C:\kpl\commodities.csv"))
            {
                string currentLine;

                // currentLine will be null when the StreamReader reaches the end of file
                while ((currentLine = sr.ReadLine()) != null)
                {
                    string[] coulumns = currentLine.Split(new char[] { ';' });
                    if (coulumns.Length > 1)
                    {
                        string commValue = coulumns[0];
                        string id = coulumns[1];

                        if (!commodities.ContainsKey(commValue))
                        {
                            //sb.AppendLine(string.Format("unknown KONZ_COMMODITY:{0}", commValue));
                            continue;
                        }


                        Comm comm;
                        if (values.ContainsKey(id))
                        {
                            comm = values[id];
                        }
                        else
                        {
                            comm = new Comm();
                            values.Add(id, comm);
                        }

                        comm.Id = id;
                        comm.Translation = commodities[commValue];

                    }
                }
            }


            using (StreamReader sr = new StreamReader(@"C:\kpl\sourcecsv28Group.csv"))
            {
                string currentLine;

                // currentLine will be null when the StreamReader reaches the end of file
                while ((currentLine = sr.ReadLine()) != null)
                {
                    if (!matGroup.Contains(currentLine))
                    {
                        matGroup.Add(currentLine);
                    }
                }
            }

            foreach (string item in matGroup)
            {

                if (!values.ContainsKey(item))
                {
                    string commodityLine = string.Format("group:{0} not found in commodity file", item);

                    sb.AppendLine(commodityLine);
                }

            }

            File.WriteAllText(@"C:\kpl\missingCommodities.txt", sb.ToString());
        }

        private static void ReadCsvSearchDocs()
        {
            Console.WriteLine("Generate csv report");
            Console.WriteLine("Source environment url: ");
            string sourceEnv = Console.ReadLine();//@"C:\kpl\SearchDocId.csv"

            Console.WriteLine("Source file path: ");
            string source = Console.ReadLine();//@"C:\kpl\SearchDocId.csv"

            Console.WriteLine("Target file path: ");
            string target = Console.ReadLine();//@"C:\kpl\SearchDocIdReport.csv"

            //try {
            using (StreamReader sr = new StreamReader(source))
            {
                // currentLine will be null when the StreamReader reaches the end of file

                List<SinequaProfile> profiles = new List<SinequaProfile>();
                Dictionary<string, List<SinequaDcoument>> queryDocs = new Dictionary<string, List<SinequaDcoument>>();
                string currentLine;
                int i = 0;
                while ((currentLine = sr.ReadLine()) != null)
                {
                    if (i == 0)
                    {
                        i++;
                        continue;
                    }

                    string[] coulumns = currentLine.Split(new char[] { ',' });
                    string profile = CleanInput(coulumns[0]);
                    SinequaProfile profileItem = profiles.FirstOrDefault(x => x.Title.Equals(profile, StringComparison.InvariantCultureIgnoreCase));
                    if (profileItem == null)
                    {
                        profileItem = new SinequaProfile();
                        profileItem.Title = profile;
                        profiles.Add(profileItem);
                    }

                    string eventType = CleanInput(coulumns[4]);
                    string resultId = CleanInput(coulumns[2]);
                    if (eventType.Equals(SineqaEventType.SearchText))
                    {
                        string queryText = CleanInput(coulumns[1]);
                        SinequaSearch searchItem = profileItem.SearchItems.FirstOrDefault(x => x.ResultId.Equals(resultId, StringComparison.InvariantCultureIgnoreCase));
                        if (searchItem == null)
                        {

                            searchItem = new SinequaSearch();
                            searchItem.QueryText = queryText;
                            searchItem.ResultId = resultId;
                            searchItem.ItemCount = 1;
                            profileItem.SearchItems.Add(searchItem);
                        }
                    }
                    else if (eventType.Equals(SineqaEventType.SearchResultLink) || eventType.Equals(SineqaEventType.DocPreview))
                    {

                        List<SinequaDcoument> documents;
                        if (!queryDocs.ContainsKey(resultId))
                        {
                            documents = new List<SinequaDcoument>();
                            queryDocs.Add(resultId, documents);
                        }
                        else
                        {
                            documents = queryDocs[resultId];
                        }


                        string documentId = CleanInput(coulumns[3]);
                        SinequaDcoument documentItem = documents.FirstOrDefault(x => x.DocId.Equals(documentId, StringComparison.InvariantCultureIgnoreCase));
                        if (documentItem == null)
                        {

                            documentItem = new SinequaDcoument();
                            documentItem.DocId = documentId;
                            documentItem.ResultId = resultId;
                            documentItem.Url = string.Format("{0}/docresult?id={1}", sourceEnv, documentId);
                            documentItem.ItemCount = 1;
                            documents.Add(documentItem);
                        }
                        else
                        {
                            documentItem.ItemCount = ++documentItem.ItemCount;
                        }

                    }

                    i++;
                }

                // group by text query and documents
                foreach (SinequaProfile profileItem in profiles)
                {
                    if (profileItem.SearchItems.Count() == 0)
                    {
                        continue;
                    }

                    List<SinequaSearch> groupedSarchItems = new List<SinequaSearch>();
                    foreach (SinequaSearch searchItem in profileItem.SearchItems)
                    {
                        if (groupedSarchItems.FirstOrDefault(x => x.QueryText.Equals(searchItem.QueryText, StringComparison.InvariantCultureIgnoreCase)) != null)
                        {
                            continue;
                        }


                        List<SinequaSearch> groupedByText = profileItem.SearchItems.FindAll(x => x.QueryText.Equals(searchItem.QueryText, StringComparison.InvariantCultureIgnoreCase));
                        if (groupedByText.Count() == 1)
                        {
                            groupedSarchItems.Add(searchItem);

                            continue;
                        }

                        SinequaSearch searchItemCumulated = new SinequaSearch();
                        searchItemCumulated.QueryText = searchItem.QueryText;
                        searchItemCumulated.ItemCount = groupedByText.Count();
                        foreach (SinequaSearch searchItemFound in groupedByText)
                        {
                            if (!queryDocs.ContainsKey(searchItemFound.ResultId))
                            {
                                continue;
                            }
                            foreach (SinequaDcoument doc in queryDocs[searchItemFound.ResultId])
                            {
                                SinequaDcoument docCheck = searchItemCumulated.DocumentItems.FirstOrDefault(x => x.DocId.Equals(doc.DocId, StringComparison.InvariantCultureIgnoreCase));
                                if (docCheck == null)
                                {
                                    searchItemCumulated.DocumentItems.Add(doc);
                                }
                                else
                                {
                                    docCheck.ItemCount = docCheck.ItemCount + doc.ItemCount;
                                }
                            }
                        }

                        groupedSarchItems.Add(searchItemCumulated);

                    }

                    profileItem.SearchItems.Clear();
                    profileItem.SearchItems.AddRange(groupedSarchItems);

                }


                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Profil;Suchbegriff;Count;DocId;Count doc;Doc URL");
                foreach (SinequaProfile profileItem in profiles)
                {
                    foreach (SinequaSearch searchItem in profileItem.SearchItems.OrderByDescending(x => x.ItemCount))
                    {
                        string commodityLine = string.Empty;

                        // query text item
                        commodityLine = string.Format("{0};{1};{2};;", profileItem.Title, searchItem.QueryText, searchItem.ItemCount);
                        sb.AppendLine(commodityLine);

                        // add document items
                        foreach (SinequaDcoument docItem in searchItem.DocumentItems.OrderByDescending(x => x.ItemCount))
                        {
                            commodityLine = string.Empty;

                            // query text item
                            commodityLine = string.Format(";;;{0};{1};{2}", docItem.DocId, docItem.ItemCount, docItem.Url);
                            sb.AppendLine(commodityLine);
                        }
                    }


                }

                File.WriteAllText(target, sb.ToString());
                Console.WriteLine("Report done");
            }

            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine("error:" + e.Message);
            //}
            //finally
            //{
            //    Console.WriteLine("press any key");
            //    Console.ReadKey();
            //}
        }

        static string CleanInput(string strIn)
        {
            // Replace invalid characters with empty strings.
            try
            {
                return strIn.Replace("\"", string.Empty);
            }
            // If we timeout when replacing invalid characters, 
            // we should return Empty.
            catch (RegexMatchTimeoutException)
            {
                return string.Empty;
            }
        }

        private class LdbItem
        {
            public string Duns { get; set; }
            public string Name { get; set; }
            public string HqDuns { get; set; }
            public string HqName { get; set; }
            public string GuDuns { get; set; }
            public string GuName { get; set; }
            public bool OutOfBusiness { get; set; }
        }

        [Serializable]
        private class OrgTreeItem: ICloneable
        {
            public string duns { get; set; }

            public string label { get; set; }

            public bool selected { get; set; }
            public bool expanded { get; set; }

            public int count { get; set; }

            public List<OrgTreeItem> items { get; } = new List<OrgTreeItem>();

            public object Clone()
            {
                using (var ms = new MemoryStream())
                {
                    var formatter = new BinaryFormatter();
                    formatter.Serialize(ms, this);
                    ms.Position = 0;

                    return (OrgTreeItem)formatter.Deserialize(ms);
                }
                //return this.MemberwiseClone();
            }
        }
    }
}
