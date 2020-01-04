using Microsoft.Office.Tools.Ribbon;
using System.Collections.ObjectModel;
using OutlookAddIn.Shared;
using System.Net.Http.Headers;
using OutlookAddIn.Classes;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            #region create organisations for dropdown
            //Get Organisations from Textfile
            OrganisationHandler OrgHandler = new OrganisationHandler();
            ObservableCollection<string> OrganisationList = new ObservableCollection<string>();
            ObservableCollection<string> OrganisationListMail = new ObservableCollection<string>();
            OrgHandler.createFile();
            OrganisationList = OrgHandler.GetOrganisations();
            OrganisationListMail = OrgHandler.GetOrganisations();

            RibbonDropDown DropdownOrganisations = Globals.Ribbons.RibbonMenu.dropDownOrg;
            RibbonDropDown DropdownOrganisationsEmail = Globals.Ribbons.RibbonMail.dropDownOrg;
            

            foreach (string Organisation in OrganisationList)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = Organisation;
                DropdownOrganisations.Items.Add(item);
            }

            foreach (string Organisation in OrganisationListMail)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = Organisation;
                DropdownOrganisationsEmail.Items.Add(item);
            }
            #endregion
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    muss ausgeführt werden, wenn Outlook heruntergefahren wird. Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
