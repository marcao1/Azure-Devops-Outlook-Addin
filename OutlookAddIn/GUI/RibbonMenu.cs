using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Net.Http.Headers;
using System.Windows;
using Microsoft.Office.Tools.Ribbon;
using OutlookAddIn.Classes;
using OutlookAddIn.GUI;
using OutlookAddIn.Shared;
using OutlookAddIn.Shared.VM;

namespace OutlookAddIn
{
    public partial class RibbonMenu
    {      
        private void RibbonMenu_Load(object sender, RibbonUIEventArgs e)
        {
            Functions.disableBtn();
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //Establish Connection to Azure DevOps
            Connection con = new Connection();
            AuthenticationHeaderValue head = con.ConnectMethod();
            Connection.bearerAuthHeader = head;

            if (head.Parameter != null)
            {
                Functions.enableBtn();
                Functions.ListProjectsMenu(con, head);
            } else
            {
                MessageBox.Show("Could not connect to Azure", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void dropDownOrg_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Connection con = new Connection();
            AuthenticationHeaderValue head = Connection.bearerAuthHeader;

            Functions.ListProjectsMenu(con, head);
        }

        private void editOrgBtn_Click(object sender, RibbonControlEventArgs e)
        {
            //Edit the txt file for organisations
            ProcessStartInfo info = new ProcessStartInfo(@"C:\Users\Public\Organisation.txt");
            Process.Start(info);
        }

        private void AddBtn_Click(object sender, RibbonControlEventArgs e)
        {
            //Application.LoadComponent(new Uri("pack://application:,,,/OutlookAddIn;UserControl/GUI/AddItem.xaml", UriKind.Relative));
            //AddItem view = new AddItem();
            //view.Show();            
        }
    }      
    }

