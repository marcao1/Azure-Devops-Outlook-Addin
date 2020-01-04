using Microsoft.Office.Tools.Ribbon;
using OutlookAddIn.Shared;
using OutlookAddIn.Shared.VM;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace OutlookAddIn.Classes
{
    public class Functions
    {
        //falls kein projekt, dann meldung schicken, boardcolumns kann man eigentlich entfernen
        //button für board einfügen (/{organisation}/board..)
        public static void ListProjectsMenu(Connection con, AuthenticationHeaderValue head)
        {
            //Lists for Items recieved from AzureDevOps
            ObservableCollection<ProjectVM> ProjectList = new ObservableCollection<ProjectVM>();
            RibbonDropDown DropdownOrganisations = Globals.Ribbons.RibbonMenu.dropDownOrg;

            //Initalize Dropdowns in Ribbons        
            RibbonDropDown DropDownProjects = Globals.Ribbons.RibbonMenu.dropDownProj;
            RibbonDropDown DropdownWorkItems = Globals.Ribbons.RibbonMenu.dropDownType;
            RibbonDropDown DropdownBoardColumns = Globals.Ribbons.RibbonMenu.dropDownCol;

            //Get Data from AzureDevOps
            try
            {
                ProjectList = con.GetProjects(DropdownOrganisations.SelectedItem.ToString(), head);
                Functions.enableBtn();  
            }
            catch
            {
                DropDownProjects.Enabled = false;
                MessageBox.Show("The Organisation is not existing", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            #region Project Dropdown Ribbons            
            DropDownProjects.Items.Clear();

            //Fill Ribbon DropDowns with Data 
            foreach (var Project in ProjectList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = Project.name;
                DropDownProjects.Items.Add(item);
            }
            #endregion

            //Lists for Items recieved from AzureDevOps
            ObservableCollection<BoardColumnVM> BoardColumnList = new ObservableCollection<BoardColumnVM>();
            ObservableCollection<WorkItemVM> WorkItemList = new ObservableCollection<WorkItemVM>();
            
            try
            {
                WorkItemList = con.GetWorkItems(head, DropdownOrganisations.SelectedItem.ToString(), DropDownProjects.SelectedItem.ToString());
                BoardColumnList = con.GetBoardColumns(DropdownOrganisations.SelectedItem.ToString(), DropDownProjects.SelectedItem.ToString());
                Functions.enableBtn();
            }
            catch
            {
                DropDownProjects.Items.Clear();
                DropdownBoardColumns.Enabled = false;
                DropdownWorkItems.Enabled = false;
                Globals.Ribbons.RibbonMenu.AddBtn.Enabled = false;
                MessageBox.Show("The Org is not existing", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            #region Type, Columns RibbonDropdowns              
            if (BoardColumnList != null)
            {
                //Fill Ribbon DropDowns with Data 
                foreach (var Column in BoardColumnList)
                {
                    RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                    item.Label = Column.name;
                    DropdownBoardColumns.Items.Add(item);
                }
            }
            if(WorkItemList != null) { 
                foreach (var WorkItem in WorkItemList)
                {
                    RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                    item.Label = WorkItem.name;
                    DropdownWorkItems.Items.Add(item);
                }
            }
            #endregion
        }

        public static void ListProjectsMail(Connection con, AuthenticationHeaderValue head)
        {
            //Lists for Items recieved from AzureDevOps
            ObservableCollection<ProjectVM> ProjectList = new ObservableCollection<ProjectVM>();
            //RibbonDropDownItem DropdownOrganisationsMenu = Globals.Ribbons.RibbonMenu.dropDownOrg.SelectedItem;
            RibbonDropDown DropdownOrganisationsMail = Globals.Ribbons.RibbonMail.dropDownOrg;
            //DropdownOrganisationsEmail.SelectedItem = DropdownOrganisationsMenu;
            
                //Get Data from AzureDevOps
                try
                {
                    ProjectList = con.GetProjects(DropdownOrganisationsMail.SelectedItem.ToString(), head);
                    Functions.enableBtn();
                }
                catch
                {
                    MessageBox.Show("The Organisation is not existing", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                }           

                #region Project Dropdown Ribbons           
                RibbonDropDown DropdownProjectsMail = Globals.Ribbons.RibbonMail.dropDownProj;

                foreach (var Project in ProjectList)
                {
                    RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                    item.Label = Project.name;
                    DropdownProjectsMail.Items.Add(item);
                }
                #endregion

                //Lists for Items recieved from AzureDevOps
                ObservableCollection<BoardColumnVM> BoardColumnList = new ObservableCollection<BoardColumnVM>();
                ObservableCollection<WorkItemVM> WorkItemListMail = new ObservableCollection<WorkItemVM>();
                //Initalize Dropdowns in Ribbons            
                RibbonDropDown DropdownWorkItemsMail = Globals.Ribbons.RibbonMail.dropDownType;
                RibbonDropDown DropdownBoardColumns = Globals.Ribbons.RibbonMail.dropDownCol;
                try
                {
                    WorkItemListMail = con.GetWorkItems(head, DropdownOrganisationsMail.SelectedItem.ToString(), DropdownProjectsMail.SelectedItem.ToString());
                    BoardColumnList = con.GetBoardColumns(DropdownOrganisationsMail.SelectedItem.ToString(), DropdownProjectsMail.SelectedItem.ToString());
                    Functions.enableBtn();
                }
                catch
                {
                    DropdownWorkItemsMail.Enabled = false;
                    DropdownBoardColumns.Enabled = false;
                    Globals.Ribbons.RibbonMail.addItemBtn.Enabled = false;
                    MessageBox.Show("The Organisation is not existing", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                #region Type, Columns RibbonDropdowns                  
                if (BoardColumnList != null)
                {
                //Fill Ribbon DropDowns with Data 
                    foreach (var Column in BoardColumnList)
                    {
                        RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                        item.Label = Column.name;
                        DropdownBoardColumns.Items.Add(item);
                    }
                }   
                if (WorkItemListMail != null)
                {
                    foreach (var WorkItem in WorkItemListMail)
                    {
                        RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                        item.Label = WorkItem.name;
                        DropdownWorkItemsMail.Items.Add(item);
                    }
                }
                #endregion
            }

        public static void disableBtn()
        {
            Globals.Ribbons.RibbonMenu.dropDownOrg.Enabled = false;
            Globals.Ribbons.RibbonMenu.dropDownProj.Enabled = false;
            Globals.Ribbons.RibbonMenu.dropDownType.Enabled = false;
            Globals.Ribbons.RibbonMenu.dropDownCol.Enabled = false;
            Globals.Ribbons.RibbonMenu.AddBtn.Enabled = false;
            Globals.Ribbons.RibbonMail.addItemBtn.Enabled = false;
            Globals.Ribbons.RibbonMail.dropDownOrg.Enabled = false;
            Globals.Ribbons.RibbonMail.dropDownProj.Enabled = false;
            Globals.Ribbons.RibbonMail.dropDownType.Enabled = false;
            Globals.Ribbons.RibbonMail.dropDownCol.Enabled = false;
        }

        public static void enableBtn()
        {
            Globals.Ribbons.RibbonMenu.dropDownOrg.Enabled = true;
            Globals.Ribbons.RibbonMenu.dropDownProj.Enabled = true;
            Globals.Ribbons.RibbonMenu.dropDownType.Enabled = true;
            Globals.Ribbons.RibbonMenu.dropDownCol.Enabled = true;
            Globals.Ribbons.RibbonMenu.AddBtn.Enabled = true;
            Globals.Ribbons.RibbonMail.addItemBtn.Enabled = true;
            Globals.Ribbons.RibbonMail.dropDownOrg.Enabled = true;
            Globals.Ribbons.RibbonMail.dropDownProj.Enabled = true;
            Globals.Ribbons.RibbonMail.dropDownType.Enabled = true;
            Globals.Ribbons.RibbonMail.dropDownCol.Enabled = true;
        }

    }
    }

