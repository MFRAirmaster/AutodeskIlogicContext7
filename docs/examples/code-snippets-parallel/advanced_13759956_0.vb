' Title: Enable Add-in in Inventor Nesting
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/enable-add-in-in-inventor-nesting/td-p/13759956
' Category: advanced
' Scraped: 2025-10-09T08:52:36.078857

using Autodesk.InventorAddins.Common;
using Inventor;
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace NestingStudyRenamer
{
    /// <summary>
    /// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
    /// that all Inventor AddIns are required to implement. The communication between Inventor and
    /// the AddIn is via the methods on this interface.
    /// </summary>
    [GuidAttribute("72571078-0228-45ef-9f4a-8dda3af5f612")]
    public class StandardAddInServer : Inventor.ApplicationAddInServer
    {

        // Inventor application object.
        private Inventor.Application m_inventorApplication;
        private Inventor.ButtonDefinition m_helloButton;
        public StandardAddInServer()
        {

        }

        #region ApplicationAddInServer Members

        public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            try
            {
                // This method is called by Inventor when it loads the addin.
                // The AddInSiteObject provides access to the Inventor Application object.
                // The FirstTime flag indicates if the addin is loaded for the first time.

                // Initialize AddIn members.
                m_inventorApplication = addInSiteObject.Application;
                m_inventorApplication.UserInterfaceManager.UserInterfaceEvents.OnEnvironmentChange += new Inventor.UserInterfaceEventsSink_OnEnvironmentChangeEventHandler(OnEnvChange);

                // TODO: Add ApplicationAddInServer.Activate implementation.
                // e.g. event initialization, command creation etc.
                // Create icon (optional, using built-in icon for now)
                // Create UI Button
                var controlDefs = m_inventorApplication.CommandManager.ControlDefinitions;

                // Load icons from embedded resource or disk
                System.Drawing.Image icon16 = PictureDispConverter.LoadResourcesImage(InvAddIn.Properties.Resources.autocad_2017_badge_16x16);
                System.Drawing.Image icon32 = PictureDispConverter.LoadResourcesImage(InvAddIn.Properties.Resources.autocad_2017_badge_32x32);

                // Convert to IPictureDisp
                stdole.IPictureDisp icon16Disp = PictureDispConverter.ToIPictureDisp(icon16);
                stdole.IPictureDisp icon32Disp = PictureDispConverter.ToIPictureDisp(icon32);

                m_helloButton = controlDefs.AddButtonDefinition(
                    "Say Hello", "SayHelloBtn",
                    CommandTypesEnum.kNonShapeEditCmdType | CommandTypesEnum.kEditMaskCmdType, "{72571078-0228-45ef-9f4a-8dda3af5f612}",
                    "Displays Hello World",
                    "Click to show message",
                    icon16Disp,  // 16x16 icon
                    icon32Disp  // 32x32 icon

                );

                //Use proper event delegate
                m_helloButton.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(HelloButton_OnExecute);
                m_helloButton.Enabled = true;

                // Add button to UI (Part Ribbon Tab)
                Ribbon ribbon = m_inventorApplication.UserInterfaceManager.Ribbons["Assembly"];
                RibbonTab tab = ribbon.RibbonTabs["id_TabTools"];
                RibbonPanel panel = tab.RibbonPanels.Add("Hello Panel",
                    "HelloPanel",
                    "{72571078-0228-45ef-9f4a-8dda3af5f612}",
                    "", false);

                panel.CommandControls.AddButton(m_helloButton, true);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void OnEnvChange(Inventor.Environment environment,
                                    EnvironmentStateEnum environmentState,
                                    EventTimingEnum beforeOrAfter,
                                    NameValueMap context,
                                    out HandlingCodeEnum handlingCode)
        {

            // Only proceed after environment change
            if (beforeOrAfter == Inventor.EventTimingEnum.kAfter)
            {

                // Example: enable button only in Nesting environment
                bool inNesting = environment.DisplayName.Equals("Nesting", StringComparison.OrdinalIgnoreCase);
                bool isAssemblyDocument = m_inventorApplication.ActiveDocument?.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject;

                if (inNesting && isAssemblyDocument)
                {
                    m_helloButton.OnExecute += new ButtonDefinitionSink_OnExecuteEventHandler(HelloButton_OnExecute);
                    m_helloButton.Enabled = inNesting;
                }

                //MessageBox.Show(m_helloButton.Enabled.ToString());

                handlingCode = HandlingCodeEnum.kEventHandled;
            }
            else
            {
                handlingCode = HandlingCodeEnum.kEventNotHandled;
            }

        }

        //Define the handler method
        virtual protected void HelloButton_OnExecute(NameValueMap Context)
        {
            InvAddIn.frmMain frm = new InvAddIn.frmMain(m_inventorApplication);

            if (frm != null)
            {
                frm.Activate();
                frm.TopMost = true;
                frm.ShowInTaskbar = false;
                frm.Show();
            }
        }

        public void Deactivate()
        {
            // This method is called by Inventor when the AddIn is unloaded.
            // The AddIn will be unloaded either manually by the user or
            // when the Inventor session is terminated

            // TODO: Add ApplicationAddInServer.Deactivate implementation
            Marshal.ReleaseComObject(m_inventorApplication);
            // Release objects.
            m_inventorApplication = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void ExecuteCommand(int commandID)
        {
            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
        }

        public object Automation
        {
            // This property is provided to allow the AddIn to expose an API 
            // of its own to other programs. Typically, this  would be done by
            // implementing the AddIn's API interface in a class and returning 
            // that class object through this property.

            get
            {
                // TODO: Add ApplicationAddInServer.Automation getter implementation
                return null;
            }
        }

        #endregion

    }
}