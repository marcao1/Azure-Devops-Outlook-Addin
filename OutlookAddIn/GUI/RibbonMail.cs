using System.Collections.ObjectModel;
using System.Net.Http.Headers;
using Microsoft.Office.Tools.Ribbon;
using OutlookAddIn.Classes;
using OutlookAddIn.Shared;
using OutlookAddIn.Shared.VM;

namespace OutlookAddIn
{
    public partial class RibbonMail
    {
        private void RibbonMail_Load(object sender, RibbonUIEventArgs e)
        {
            //Establish Connection to Azure DevOps
            Connection con = new Connection();
            AuthenticationHeaderValue head = Connection.bearerAuthHeader;

            if (head.Parameter != null)
            {
                Functions.ListProjectsMail(con, head);
            }
        }
       
        private void dropDownOrg_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Connection con = new Connection();
            AuthenticationHeaderValue head = Connection.bearerAuthHeader;

            if (head.Parameter != null)
            {
                Functions.ListProjectsMail(con, head);
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            //muss noch geändert werden
            Connection con = new Connection();
            AuthenticationHeaderValue head = Connection.bearerAuthHeader;
            con.CreateWorkItem(head, dropDownType.SelectedItem.ToString(), dropDownOrg.SelectedItem.ToString(), dropDownProj.SelectedItem.ToString(), "test");
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
