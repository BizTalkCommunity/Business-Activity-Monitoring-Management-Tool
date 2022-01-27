using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Win32;
using System.IO;

namespace BizTalk.BAM.Management.Tool
{
    public partial class BAMMgmtUtlForm : Form
    {
        #region Properties

        private string bamTrackingPath = string.Empty;
        private const bool keepOldLogText = true;
        private const bool overwriteOldLogText = false;
        private string bamPrimaryImportDB = string.Empty;
        private string server = string.Empty;
        private TreeNode activitiesSelectedNode = null;

        private enum ItemType
        {
            Activity,
            Views,
            Alerts,
            Accounts,
            Config
        }

        private enum NodeType
        {
            View,
            Alert,
            Account
        }

        #endregion

        #region Constructor

        public BAMMgmtUtlForm()
        {
            InitializeComponent();

            bamPrimaryImportDB = BAM.Default.Db;
            if (BAM.Default.Server != string.Empty)
                server = BAM.Default.Server;
            else server = System.Environment.MachineName;
            
            bamPrimaryImportDBLbl.Text = bamPrimaryImportDB;
            serverLbl.Text = server;
            bamDefFileTextBox.Text = BAM.Default.InitialBAMDefinition;
            bamTrackingPath = BAM.Default.BMPath;
            BAMConfigFileTextBox.Text = BAM.Default.InitialFolder + "MyConfig.xml";
            FillItemsListTreeView(overwriteOldLogText);

            RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\BizTalk Server\3.0");
            string value = Convert.ToString(key.GetValue("ProductName"));
            labelversion.Text = value;
        }

        #endregion

        #region Private methods

        private string ExecuteCommand(string cmd, object sender, bool keepOldLogText)
        {
            String finalOut = string.Empty;

            try
            {
                this.Cursor = Cursors.WaitCursor;

                if (sender != null && sender is Button)
                {
                    Button btn = (Button)sender;
                    btn.Enabled = false;
                }

                // Check if bm.exe (out-of-the-box command line tool) exists
                System.IO.FileInfo fi = new System.IO.FileInfo(bamTrackingPath);
                if (!fi.Exists)
                {
                    throw new Exception(string.Format("Could not find bm.exe at path '{0}'. Please change BMPath in the config file!", bamTrackingPath));
                }

                // Getting the encoding from registry 
                int oemCodePage = 0;
                try
                {
                    RegistryKey codepageKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Nls\CodePage");
                    oemCodePage = int.Parse((string)codepageKey.GetValue("OEMCP"));
                }
                catch
                {}

                TraceResultsInTool(string.Format("Executing command: {0} {1}", bamTrackingPath, cmd), keepOldLogText);
                string bamDefFile = bamDefFileTextBox.Text.Trim();
                Process bmProcess = new Process();
                ProcessStartInfo bsi = new ProcessStartInfo(bamTrackingPath, cmd);
                bsi.RedirectStandardOutput = true;
                bsi.UseShellExecute = false;
                bsi.CreateNoWindow = true;
                bmProcess.StartInfo = bsi;
                bmProcess.Start();

                if (oemCodePage == 0)
                {
                    finalOut = bmProcess.StandardOutput.ReadToEnd();
                }
                else
                {
                    StreamReader newStandardOutput = new StreamReader(bmProcess.StandardOutput.BaseStream, Encoding.GetEncoding(oemCodePage));
                    finalOut = newStandardOutput.ReadToEnd();
                }

                bmProcess.WaitForExit();
                if (bmProcess.HasExited)
                    TraceResultsInTool(finalOut, true);
                else throw new Exception("Generic error");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sender != null && sender is Button)
                {
                    Button btn = (Button)sender;
                    btn.Enabled = true;
                }

                this.Cursor = Cursors.Default;
            }

            if (finalOut.Contains("ERROR: "))
            {
                MessageBox.Show(finalOut.Substring(finalOut.IndexOf("ERROR: ")), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return finalOut;
        }

        private void SelectItemType(ItemType itemType)
        {
            itemsListTreeView.Tag = itemType;

            switch (itemType)
            {
                case ItemType.Activity:
                    listGroupBox.Text = "Activities";
                    break;
                case ItemType.Views:
                    listGroupBox.Text = "Views";
                    break;
                case ItemType.Alerts:
                    listGroupBox.Text = "Alerts";
                    break;
                case ItemType.Accounts:
                    listGroupBox.Text = "Accounts";
                    break;
                case ItemType.Config:
                    listGroupBox.Text = "Config";
                    break;
            }
        }

        private void FillItemsListTreeView(bool KeepTheOldLogText)
        {
            itemsListTreeView.Nodes.Clear();

            string tmp;
            int index;
            string[] items;

            switch (itemsTabControl.SelectedTab.Name)
            {
               case "activitiesTabPage":
                    activityNameTextBox.Text = string.Empty;
                    SelectItemType(ItemType.Activity);
                    tmp = ExecuteCommand(string.Format("get-activities -Server:\"{0}\" -Database:\"{1}\"", server, bamPrimaryImportDB), null, KeepTheOldLogText);
                    index = tmp.IndexOf("All activities in the database:\r\n");
                    if (index > 0)
                    {
                        tmp = tmp.Substring(index + 33);
                        items = tmp.Split('\r');
                        foreach (string activity in items)
                        {
                            if (activity != "\n" && activity != "\r" && activity != "\r\n")
                                itemsListTreeView.Nodes.Add(activity.Replace("\n", string.Empty));
                        }
                    }
                    else
                    {
                        itemsListTreeView.Nodes.Add("No Activities are defined");
                    }
                    break;

                case "viewsTabPage":
                    viewNameTextBox.Text = string.Empty;
                    SelectItemType(ItemType.Views);
                    tmp = ExecuteCommand(string.Format("get-views -Server:\"{0}\" -Database:\"{1}\"", server, bamPrimaryImportDB), null, KeepTheOldLogText);
                    index = tmp.IndexOf("All views in the database:\r\n");
                    if (index > 0)
                    {
                        tmp = tmp.Substring(index + 28);
                        items = tmp.Split('\r');
                        foreach (string view in items)
                        {
                            if (view != "\n" && view != "\r" && view != "\r\n")
                                itemsListTreeView.Nodes.Add(view.Replace("\n", string.Empty));
                        }
                     }
                     else
                     {
                         itemsListTreeView.Nodes.Add("No views are defined");
                     }
                    break;

                case "alertsTabPage":
                    alertNameTextBox.Text = string.Empty;
                    SelectItemType(ItemType.Alerts);
                    tmp = ExecuteCommand(string.Format("get-alerts -Server:\"{0}\" -Database:\"{1}\"", server, bamPrimaryImportDB), null, KeepTheOldLogText);
                    index = tmp.IndexOf("Alerts for view ");
                    if (index > 0)
                    {
                        tmp = tmp.Substring(index);
                        items = tmp.Split('\r');
                        string lastView = string.Empty;
                        TreeNode activeNode = null;
                        foreach (string item in items)
                        {
                            if (item.Equals("\n") || item.Equals("\r") || item.Equals("\r\n"))
                                continue;
                            if (item.IndexOf("Alerts for view ") > -1)
                            {
                                // Parse new view name
                                string viewName = item.Substring(item.IndexOf('\'') + 1);
                                viewName = viewName.TrimEnd(new char[] { '\'', ':' });
                                TreeNode newViewNode = new TreeNode(viewName);
                                activeNode = newViewNode;
                                newViewNode.Tag = NodeType.View;
                                itemsListTreeView.Nodes.Add(newViewNode);
                            }
                            else
                            {
                                activeNode.Nodes.Add(item.TrimStart('\n'));
                            }
                        }
                        itemsListTreeView.ExpandAll();
                    }
                    else
                    {
                        itemsListTreeView.Nodes.Add("No alerts are defined");
                    }
                    break;

                case "accountsTabPage":
                    accountsViewNameTextBox.Text = string.Empty;
                    accountsAccountNameTextBox.Text = string.Empty;
                    SelectItemType(ItemType.Accounts);
                    tmp = ExecuteCommand(string.Format("get-views -Server:\"{0}\" -Database:\"{1}\"", server, bamPrimaryImportDB), null, KeepTheOldLogText);
                    index = tmp.IndexOf("All views in the database:\r\n");
                    if (index > 0)
                    {
                        tmp = tmp.Substring(index + 28);
                        items = tmp.Split('\r');
                        foreach (string view in items)
                        {
                            if (view.Equals("\n") || view.Equals("\r") || view.Equals("\r\n"))
                                continue;

                            string viewName = view.Replace("\n", string.Empty);
                            TreeNode newNode = new TreeNode(viewName);
                            itemsListTreeView.Nodes.Add(newNode);
                            newNode.Tag = NodeType.Account;
                            tmp = ExecuteCommand(string.Format("get-accounts -View:\"{0}\" -Server:\"{1}\" -Database:\"{2}\"",viewName, server, bamPrimaryImportDB), null, keepOldLogText);
                            string accountsStr = string.Format("Accounts for view '{0}':\r\n", viewName);
                            index = tmp.IndexOf(accountsStr);
                            if (index > 0)
                            {
                                tmp = tmp.Substring(index + accountsStr.Length);
                                items = tmp.Split('\r');
                                foreach (string account in items)
                                {
                                    if (account.Equals("\n") || account.Equals("\r") || account.Equals("\r\n"))
                                        continue;

                                    newNode.Nodes.Add(account.Replace("\n", string.Empty));
                                }
                            }
                            else
                            {
                                newNode.Nodes.Add("No accounts are defined for this view");
                            }
                        }
                    }
                    else
                    {
                        itemsListTreeView.Nodes.Add("No views are defined");
                    }
                    break;

                case "configTabPage":
                    BAMConfigFileTextBox.Text = string.Empty;
                    SelectItemType(ItemType.Config);
                    break;
                
            }

            itemsListTreeView.ExpandAll();
        }

        private void TraceResultsInTool(string message)
        {
            logTextBox.Text = message;
            Application.DoEvents();
        }

        private void TraceResultsInTool(string message, bool keepOldLogText)
        {
            if (keepOldLogText)
                logTextBox.Text += Environment.NewLine + message;
            else
                logTextBox.Text = message;
        }


        #endregion

        #region Events

        private void deployBamDefBtn_Click(object sender, EventArgs e)
        {
            //if (MessageBox.Show(string.Format("Are you sure you want to deploy bam definition '{0}'?", bamDefFileTextBox.Text.Trim()), "Deploy bam definition?", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
            //   return;

            string cmd = "deploy-all -DefinitionFile:\"{0}\" -Server:\"{1}\" -Database:\"{2}\"";
            cmd = string.Format(cmd, bamDefFileTextBox.Text.Trim(), server, bamPrimaryImportDB);
            ExecuteCommand(cmd, sender, overwriteOldLogText);
            FillItemsListTreeView(keepOldLogText);
        }

        private void updateBamDefBtn_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(string.Format("Are you sure you want to update bam definition '{0}'?", bamDefFileTextBox.Text.Trim()), "Update bam definition?", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                return;

            string cmd = "update-all -DefinitionFile:\"{0}\" -Server:\"{1}\" -Database:\"{2}\"";
            cmd = string.Format(cmd, bamDefFileTextBox.Text.Trim(), server, bamPrimaryImportDB);
            ExecuteCommand(cmd, sender, overwriteOldLogText);
            FillItemsListTreeView(keepOldLogText);
        }

        private void removeBamDefBtn_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(string.Format("Are you sure you want to delete bam definition '{0}'?", bamDefFileTextBox.Text.Trim()), "Delete bam definition?", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                return;

            string cmd = "remove-all -DefinitionFile:\"{0}\" -Server:\"{1}\" -Database:\"{2}\"";
            cmd = string.Format(cmd, bamDefFileTextBox.Text.Trim(), server, bamPrimaryImportDB);
            ExecuteCommand(cmd, sender, overwriteOldLogText);
            FillItemsListTreeView(keepOldLogText);
        }

        private void browseBAMDefBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openBrowseFileDialog = new OpenFileDialog();
            openBrowseFileDialog.InitialDirectory = BAM.Default.InitialFolder;
            openBrowseFileDialog.Filter = "BizTalk BAM Definition file (*.xls)|*.xls|All files (*.*)|*.*";
            openBrowseFileDialog.FilterIndex = 1;
            openBrowseFileDialog.RestoreDirectory = true;
            openBrowseFileDialog.Title = "Open BizTalk BAM Definition file";
            if (openBrowseFileDialog.ShowDialog() == DialogResult.OK)
                bamDefFileTextBox.Text = openBrowseFileDialog.FileName;
        }

        private void executeCommandBtn_Click(object sender, EventArgs e)
        {
            ExecuteCommand(commandTextBox.Text.Trim(), sender, overwriteOldLogText);
            FillItemsListTreeView(keepOldLogText);
        }

        private void helpBtn_Click(object sender, EventArgs e)
        {
            string cmd = "help";
            ExecuteCommand(cmd, sender, overwriteOldLogText);
        }

        private void itemsTabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillItemsListTreeView(overwriteOldLogText);
        }

        private void updateItemsListBtn_Click(object sender, EventArgs e)
        {
            FillItemsListTreeView(overwriteOldLogText);
        }

        private void copyToItemTabBtn_Click(object sender, EventArgs e)
        {
            ItemType itemType = (ItemType)itemsListTreeView.Tag;
            if (itemsListTreeView.SelectedNode != null)
            {
                switch (itemType)
                {
                    case ItemType.Activity:
                        itemsTabControl.SelectedTab = activitiesTabPage;
                        activityNameTextBox.Text = itemsListTreeView.SelectedNode.Text;
                        break;
                    case ItemType.Views:
                        itemsTabControl.SelectedTab = viewsTabPage;
                        viewNameTextBox.Text = itemsListTreeView.SelectedNode.Text;
                        break;
                    case ItemType.Alerts:
                        itemsTabControl.SelectedTab = alertsTabPage;
                        if (itemsListTreeView.SelectedNode.Tag == null)
                            alertNameTextBox.Text = itemsListTreeView.SelectedNode.Parent.Text;
                        else
                            alertNameTextBox.Text = itemsListTreeView.SelectedNode.Text;
                        break;
                    case ItemType.Accounts:
                        itemsTabControl.SelectedTab = accountsTabPage;
                        if (itemsListTreeView.SelectedNode.Tag == null)
                        {
                            accountsViewNameTextBox.Text = itemsListTreeView.SelectedNode.Parent.Text;
                            accountsAccountNameTextBox.Text = itemsListTreeView.SelectedNode.Text;
                        }
                        else
                        {
                            accountsViewNameTextBox.Text = string.Empty;
                            accountsAccountNameTextBox.Text = string.Empty;
                        }
                        break;
                }
            }
        }

        private void removeActivityBtn_Click(object sender, EventArgs e)
        {
            RemoveActivity(activityNameTextBox.Text.Trim(), sender);
        }

        private void RemoveActivity(string activityName, object sender)
        {
            if (MessageBox.Show(string.Format("Are you sure you want to delete activity '{0}'?", activityName), "Delete activity?", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                return;
            string cmd = "remove-activity -Name:\"{0}\" -Server:\"{1}\" -Database:\"{2}\"";
            cmd = string.Format(cmd, activityName, server, bamPrimaryImportDB);
            ExecuteCommand(cmd, sender, overwriteOldLogText);
            FillItemsListTreeView(keepOldLogText);
        }

        private void itemsListTreeView_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                // Show context menu
                ContextMenu menu = new ContextMenu();
                MenuItem item;

                switch ((ItemType)itemsListTreeView.Tag)
                {
                    case ItemType.Activity:
                        item = new MenuItem("Remove");
                        menu.MenuItems.Add(item);
                        item.Click += new System.EventHandler(this.RemoveActivityMenuIteClick);
                        break;

                    case ItemType.Views:
                        item = new MenuItem("Remove");
                        menu.MenuItems.Add(item);
                        item.Click += new System.EventHandler(this.RemoveViewMenuIteClick);
                        break;

                    case ItemType.Alerts:
                        item = new MenuItem("Remove");
                        menu.MenuItems.Add(item);
                        item.Click += new System.EventHandler(this.RemoveAlertsMenuIteClick);
                        break;

                    default:
                        break;
                }

                if (menu.MenuItems != null && menu.MenuItems.Count > 0)
                    itemsListTreeView.ContextMenu = menu;
                else
                    itemsListTreeView.ContextMenu = null;
              }
        }

        private void RemoveActivityMenuIteClick(object sender, System.EventArgs e)
        {
            RemoveActivity(activitiesSelectedNode.Text, null);
        }

        private void RemoveViewMenuIteClick(object sender, System.EventArgs e)
        {
            RemoveView(activitiesSelectedNode.Text, null);
        }

        private void RemoveAlertsMenuIteClick(object sender, System.EventArgs e)
        {
            string selectedNodeName;
            if (activitiesSelectedNode.Tag == null)
                selectedNodeName = activitiesSelectedNode.Parent.Text;
            else
                selectedNodeName = activitiesSelectedNode.Text;

            RemoveAlert(selectedNodeName, null);
        }
        
        private void forcedRemoveActivityBtn_Click(object sender, EventArgs e)
        {
            // Remove associated views/alerts first (more to remove first? accounts etc?)
        }

        private void removeViewBtn_Click(object sender, EventArgs e)
        {
            RemoveView(viewNameTextBox.Text.Trim(), sender);
        }

        private void RemoveView(string viewName, object sender)
        {
            if (MessageBox.Show(string.Format("Are you sure you want to delete view '{0}'?", viewName), "Delete view?", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                return;
            string cmd = "remove-view -Name:\"{0}\" -Server:\"{1}\" -Database:\"{2}\"";
            cmd = string.Format(cmd, viewName, server, bamPrimaryImportDB);
            ExecuteCommand(cmd, sender, overwriteOldLogText);
            FillItemsListTreeView(keepOldLogText);
        }

        private void removeAlertBtn_Click(object sender, EventArgs e)
        {
            RemoveAlert(alertNameTextBox.Text.Trim(), sender);
        }

        private void RemoveAlert(string viewName, object sender)
        {
            if (MessageBox.Show(string.Format("Are you sure you want to delete all alerts for view '{0}'?", viewName), "Delete alerts?", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                return;
            string cmd = "remove-alerts -View:\"{0}\" -Server:\"{1}\" -Database:\"{2}\"";
            cmd = string.Format(cmd, viewName, server, bamPrimaryImportDB);
            ExecuteCommand(cmd, sender, overwriteOldLogText);
            FillItemsListTreeView(keepOldLogText);
        }

        private void browseBAMConfigFileBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openBrowseFileDialog = new OpenFileDialog();
            openBrowseFileDialog.InitialDirectory = BAM.Default.InitialFolder;
            openBrowseFileDialog.Filter = "BizTalk BAM Configuration Xml file (*.xml)|*.xml|All files (*.*)|*.*";
            openBrowseFileDialog.FilterIndex = 1;
            openBrowseFileDialog.RestoreDirectory = true;
            openBrowseFileDialog.Title = "Open BizTalk BAM Configuration Xml file";
            if (openBrowseFileDialog.ShowDialog() == DialogResult.OK)
                BAMConfigFileTextBox.Text = openBrowseFileDialog.FileName;
        }

        private void exportBAMConfigFileBtn_Click(object sender, EventArgs e)
        {
            string cmd = "get-config -Filename:\"{0}\" -Server:\"{1}\" -Database:\"{2}\"";
            cmd = string.Format(cmd, BAMConfigFileTextBox.Text.Trim(), server, bamPrimaryImportDB);
            ExecuteCommand(cmd, sender, overwriteOldLogText);
        }

        private void updateBAMConfigFileBtn_Click(object sender, EventArgs e)
        {
            string cmd = "update-config -Filename:\"{0}\"";
            cmd = string.Format(cmd, BAMConfigFileTextBox.Text.Trim());
            ExecuteCommand(cmd, sender, overwriteOldLogText);
        }

        private void addAccountBtn_Click(object sender, EventArgs e)
        {
            string cmd = string.Format("add-account -AccountName:\"{0}\" -View:\"{1}\" -Server:\"{2}\" -Database:\"{3}\"", accountsAccountNameTextBox.Text.Trim(), accountsViewNameTextBox.Text.Trim(), server, bamPrimaryImportDB);
            cmd = string.Format(cmd, bamDefFileTextBox.Text.Trim());
            ExecuteCommand(cmd, sender, overwriteOldLogText);
            FillItemsListTreeView(keepOldLogText);
        }

        private void removeAccountBtn_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(string.Format("Are you sure you want to remove account '{0}' from view '{1}'?", accountsAccountNameTextBox.Text.Trim(), accountsViewNameTextBox.Text.Trim()), "Remove account?", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
                return;

            string cmd = string.Format("remove-account -AccountName:\"{0}\" -View:\"{1}\" -Server:\"{2}\" -Database:\"{3}\"", accountsAccountNameTextBox.Text.Trim(), accountsViewNameTextBox.Text.Trim(), server, bamPrimaryImportDB);
            cmd = string.Format(cmd, bamDefFileTextBox.Text.Trim());
            ExecuteCommand(cmd, sender, overwriteOldLogText);
            FillItemsListTreeView(keepOldLogText);
        }

        private void updateLivedataworkbookBtn_Click(object sender, EventArgs e)
        {
            string cmd = string.Format("update-livedataworkbook -Name:\"{0}\" -Server:\"{1}\" -Database:\"{2}\"", bamDefFileTextBox.Text, server, bamPrimaryImportDB);
            ExecuteCommand(cmd, sender, overwriteOldLogText);
        }

        private void regenerateLivedataworkbookBtn_Click(object sender, EventArgs e)
        {
            string cmd = string.Format("regenerate-livedataworkbook -WorkbookName:\"{0}\" -Server:\"{1}\" -Database:\"{2}\"", bamDefFileTextBox.Text, server, bamPrimaryImportDB);
            ExecuteCommand(cmd, sender, overwriteOldLogText);
        }


        private void itemsListTreeView_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            activitiesSelectedNode = e.Node;
        }

        #endregion

    }
}