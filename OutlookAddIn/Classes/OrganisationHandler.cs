using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddIn.Classes
{
    public class OrganisationHandler
    {
        public void createFile()
        {
            string path = @"C:\Users\Public\Organisation.txt";
           
            if (!System.IO.File.Exists(path))
                File.Create(path);
        }

        public ObservableCollection<string> GetOrganisations()
        {
            string path = @"C:\Users\Public\Organisation.txt";

            ObservableCollection<string> OrganisationList = readFile(path);

            return OrganisationList;
        }
        
        public static ObservableCollection<string> readFile(string FilePath)
        {
            string[] readText = File.ReadAllLines(FilePath);
            ObservableCollection<string> Organisations = new ObservableCollection<string>();

            if (readText.Length == 0)
            {
                Organisations.Add("Select..");
            }
            else
            {
                foreach (string s in readText)
                {
                    Organisations.Add(s);
                }
            }
            return Organisations;
        }

        public static void addOrganisation(string OrgName, string FilePath)
        {
            using (StreamWriter sw = File.AppendText(FilePath))
            {
                sw.WriteLine(OrgName);
            }
        }
    }
}
