using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AD_Unlocker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // List that contains data that's displayed in listbox
        List<string> adUsers = new List<string>();

        // Gets Active Directory Path
        private string GetCurrentDomainPath()
        {
            DirectoryEntry de = new DirectoryEntry("LDAP://RootDSE");

            return "LDAP://" + de.Properties["defaultNamingContext"][0].ToString();
        }

        // Adds ability to search listbox for specific value
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            listBox1.SelectedIndex = listBox1.FindString(textBox1.Text);
        }

        // When the Unlock User button is pressed 
        private void button1_Click(object sender, EventArgs e)
        {
            UserUnlocker();
        }

        // Sets Enter name: placeholder on form load and populates listbox
        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.PlaceholderText = "Enter name";
            LockedUsersQuery();
        }

        // Populates listbox with list of all locked users
        private void LockedUsersQuery()
        {
            Cursor.Current = Cursors.WaitCursor;

            adUsers.Clear();

            var searcher = new DirectorySearcher()
            {
                Filter = "(&(objectCategory=person)(objectClass=user)(lockoutTime>=1)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))",
                PageSize = 1000,
                PropertiesToLoad = { "sAMAccountName" }
            };

            using (var results = searcher.FindAll())
            {

                foreach (SearchResult result in results)
                {
                    var username = result.Properties["sAMAccountName"][0].ToString();

                    adUsers.Add(username);
                }
            }

            ResultsUpdater();

            Cursor.Current = Cursors.Default;
        }

        // Displays additional user info when Enter key is pressed in textbox
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                AdditionalInfo();
            }

        }

        // Function for unlocking user
        private void UserUnlocker()
        {
            try
            {
                string selectedUser = listBox1.SelectedItem.ToString();

                DirectoryEntry userEntry = new DirectoryEntry(GetCurrentDomainPath());

                DirectorySearcher search = new DirectorySearcher(userEntry);

                search.Filter = $"(&(objectCategory=person)(objectClass=user)(sAMAccountName={listBox1.SelectedItem}))";

                SearchResult user = search.FindOne();

                DirectoryEntry userInfo = user.GetDirectoryEntry();

                userInfo.Properties["LockOutTime"].Value = 0;

                userInfo.CommitChanges();

                userInfo.Close();

                userEntry.Close();

                LockedUsersQuery();

                MessageBox.Show("Success!", $"{selectedUser}'s account has been unlocked!");
            }
            catch
            {
                MessageBox.Show("Error!", "An Error has occured");

            }
        }

        // Updates Listbox to reflect changes
        private void ResultsUpdater()
        {
            listBox1.Items.Clear();

            foreach (var item in adUsers)
            {
                listBox1.Items.Add(item.ToString());
            }

            listBox1.SelectedItem = listBox1.Items[0];

        }


        // Pulls up additional user data
        private void listBox1_DoubleClick_1(object sender, EventArgs e)
        {
            AdditionalInfo();
        }

        // Displays users full name and email
        private void AdditionalInfo()
        {
            if (listBox1.SelectedItem != null)
            {
                try
                {

                    DirectorySearcher search = new DirectorySearcher(GetCurrentDomainPath());

                    search.Filter = $"(&(objectCategory=person)(objectClass=user)(sAMAccountName={listBox1.SelectedItem}))";

                    SearchResult result = search.FindOne();

                    if (result != null)
                    {

                        string displayName = result.Properties["displayName"][0].ToString();

                        string email = result.Properties["mail"][0].ToString();

                        var newline = System.Environment.NewLine;

                        StringBuilder message = new StringBuilder();

                        message.Append(displayName + newline + email);

                        string title = "Unlock this user?";

                        MessageBoxButtons buttons = MessageBoxButtons.YesNo;

                        DialogResult dr = MessageBox.Show(message.ToString(), title, buttons);

                        if (dr == DialogResult.Yes)
                        {
                            UserUnlocker();

                        }

                    }
                }
                catch { }
            }
        }

        // Makes Enter key display additional user information
        private void listBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                AdditionalInfo();
            }
        }

        // Allows user to navigate listbox through textbox with up/down keys
        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Up)
            {
                if (listBox1.SelectedIndex > 0)
                    this.listBox1.SelectedIndex--;
                e.Handled = true;

            }
            else if (e.KeyCode == Keys.Down )
            {
                if (listBox1.SelectedIndex < (listBox1.Items.Count - 1))
                   this.listBox1.SelectedIndex++;
                e.Handled = true;
            }
        }

        // Refreshes data when refresh button is pressed
        private void toolStripButton1_Click(object sender, EventArgs e)
        {

            LockedUsersQuery();

        }

        // Changes cursor when hovering over refresh button
        private void toolStripButton1_MouseEnter(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Hand;
        }

        private void toolStripButton1_MouseLeave(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Default;
        }
    }
}

