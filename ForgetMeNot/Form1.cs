using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;


namespace ForgetMeNot
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }


        private void ExportToExcel(DataGridView dataSourceGrid)
        {
            this.Cursor = Cursors.WaitCursor;
            Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Excel._Workbook ExcelBook;
            Excel._Worksheet ExcelSheet;

            int i = 0;
            int j = 0;

            //create object of excel
            ExcelBook = (Excel._Workbook)ExcelApp.Workbooks.Add(1);
            ExcelSheet = (Excel._Worksheet)ExcelBook.ActiveSheet;
            //export header
            for (i = 1; i <= dataSourceGrid.Columns.Count; i++)
            {
                ExcelSheet.Cells[1, i] = dataSourceGrid.Columns[i - 1].HeaderText;
            }

            //export data
            for (i = 1; i <= dataSourceGrid.RowCount; i++)
            {
                for (j = 1; j <= dataSourceGrid.Columns.Count; j++)
                {
                    ExcelSheet.Cells[i + 1, j] = dataSourceGrid.Rows[i - 1].Cells[j - 1].Value;
                }
            }

            ExcelApp.Visible = true;
            this.Cursor = Cursors.Arrow;

            //set font Khmer OS System to data range
            Excel.Range r1 = ExcelSheet.Cells[1, 1];
            Excel.Range r2 = ExcelSheet.Cells[dataSourceGrid.RowCount + 1, dataSourceGrid.Columns.Count];
            Excel.Range myRange = ExcelSheet.get_Range(r1, r2);
            Excel.Font x = myRange.Font;
            x.Name = "Arial";
            x.Size = 10;

            //set bold font to column header
            r1 = ExcelSheet.Cells[1, 1];
            r2 = ExcelSheet.Cells[1, dataSourceGrid.Columns.Count];
            myRange = ExcelSheet.get_Range(r1, r2);
            x = myRange.Font;
            x.Bold = true;
            x.Underline = true;
            //autofit all columns
            myRange.EntireColumn.AutoFit();

            //Reset Objects
            ExcelSheet = null;
            ExcelBook = null;
            ExcelApp = null;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'forgetMeNotDataSet.ForgetMeNot' table. You can move, or remove it, as needed.
            this.forgetMeNotTableAdapter.Fill(this.forgetMeNotDataSet.ForgetMeNot);
        }

        private void forgetMeNotBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.forgetMeNotBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.forgetMeNotDataSet);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(display_NameTextBox.Text) && String.IsNullOrWhiteSpace(passwordTextBox.Text) && String.IsNullOrWhiteSpace(usernameTextBox.Text))
            {
                MessageBox.Show("Please Enter at least a Display Name, username, and password.");
                return;
            }

            this.Cursor = Cursors.WaitCursor;
            if (String.IsNullOrWhiteSpace(application_IDTextBox.Text))
            {
                ParseFormDataAndAddNewRow();
            }
            else
            {
                //We are editing an item, so updated the item, based on ID in text box?
                //TODO - Figure out how to update.

                int index = FindRecordByApplicationIDAndReturnIndexInDataSet(application_IDTextBox.Text);

                if (index >= 0)
                {
                    SaveFormDataToDataSetRow(index);
                }
            }


            this.forgetMeNotBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.forgetMeNotDataSet);
            ClearEntryFields();
            this.Cursor = Cursors.Arrow;
        }

        private int FindRecordByApplicationIDAndReturnIndexInDataSet(string appID)
        {
            int index = -1;
            IEnumerable<DataRow> rows = forgetMeNotDataSet.ForgetMeNot.Rows.OfType<DataRow>();
            ForgetMeNotDataSet.ForgetMeNotRow currentRow = (ForgetMeNotDataSet.ForgetMeNotRow)rows.First(x => x["APPLICATION_ID"].ToString() == appID);

            if(currentRow != null)
                index = forgetMeNotDataSet.ForgetMeNot.Rows.IndexOf((DataRow)currentRow);

            return index;
        }

        private void SaveFormDataToDataSetRow(int index)
        {
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["Display_Name"] = display_NameTextBox.Text;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["Application_Name"] = application_NameTextBox.Text;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["Application_URL"] = application_URLTextBox.Text;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["Login_URL"] = login_URLTextBox.Text;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["Username"] = usernameTextBox.Text;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["Email"] = emailTextBox.Text;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["Password"] = passwordTextBox.Text;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["Software_Key"] = software_KeyTextBox.Text;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["CD_Key"] = cD_KeyTextBox.Text;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["General_Key_1"] = general_Key_1TextBox.Text;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["General_Key_2"] = general_Key_2TextBox.Text;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["General_Key_3"] = general_Key_3TextBox.Text;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["General_Key_4"] = general_Key_4TextBox.Text;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["General_Key_5"] = general_Key_5TextBox.Text;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["Has_Credit_Card_Info"] = has_Credit_Card_InfoCheckBox.Checked;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["Has_Address_Info"] = has_Address_InfoCheckBox.Checked;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["Has_Phone_Number"] = has_Phone_NumberCheckBox.Checked;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["Is_Auto_Renewal"] = is_Auto_RenewalCheckBox.Checked;
            forgetMeNotDataSet.ForgetMeNot.Rows[index]["Notes"] = notesTextBox.Text;
        }

        private void ParseFormDataAndAddNewRow()
        {
            string appID = application_IDTextBox.Text;
            string displayName = display_NameTextBox.Text;
            string applicationName = application_NameTextBox.Text;
            string applicationURL = application_URLTextBox.Text;
            string loginURL = login_URLTextBox.Text;
            string username = usernameTextBox.Text;
            string email = emailTextBox.Text;
            string password = passwordTextBox.Text;
            string softwareKey = software_KeyTextBox.Text;
            string cdKey = cD_KeyTextBox.Text;
            string gk1 = general_Key_1TextBox.Text;
            string gk2 = general_Key_2TextBox.Text;
            string gk3 = general_Key_3TextBox.Text;
            string gk4 = general_Key_4TextBox.Text;
            string gk5 = general_Key_5TextBox.Text;
            bool hasCreditCard = has_Credit_Card_InfoCheckBox.Checked;
            bool hasAddress = has_Address_InfoCheckBox.Checked;
            bool hasPhone = has_Phone_NumberCheckBox.Checked;
            bool isAutoRenewal = is_Auto_RenewalCheckBox.Checked;
            string notes = notesTextBox.Text;

            forgetMeNotDataSet.ForgetMeNot.AddForgetMeNotRow
                (displayName, applicationName, applicationURL, loginURL, username, email, password,
                softwareKey, cdKey, gk1, gk2, gk3, gk4, gk5, hasCreditCard,
                hasAddress, hasPhone, isAutoRenewal, notes);
        }

        private void ClearEntryFields()
        {
            application_IDTextBox.Text = String.Empty;
            display_NameTextBox.Text = String.Empty;
            application_NameTextBox.Text = String.Empty;
            application_URLTextBox.Text = String.Empty;
            login_URLTextBox.Text = String.Empty;
            usernameTextBox.Text = String.Empty;
            emailTextBox.Text = String.Empty;
            passwordTextBox.Text = String.Empty;
            software_KeyTextBox.Text = String.Empty;
            cD_KeyTextBox.Text = String.Empty;
            general_Key_1TextBox.Text = String.Empty;
            general_Key_2TextBox.Text = String.Empty;
            general_Key_3TextBox.Text = String.Empty;
            general_Key_4TextBox.Text = String.Empty;
            general_Key_5TextBox.Text = String.Empty;
            has_Credit_Card_InfoCheckBox.Checked = false;
            has_Address_InfoCheckBox.Checked = false;
            has_Phone_NumberCheckBox.Checked = false;
            is_Auto_RenewalCheckBox.Checked = false;
            notesTextBox.Text = String.Empty;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ClearEntryFields();
            DataRowView dataRowView = (DataRowView)forgetMeNotBindingSource.List[comboBox1.SelectedIndex];
            DataRow dataRow = dataRowView.Row;
            FillFormWithRowData(dataRow);
            //forgetMeNotDataSet.ForgetMeNot[comboBox1.SelectedIndex];

        }

        private void FillFormWithRowData(DataRow dataRow)
        {
            try
            {
                application_IDTextBox.Text = dataRow["APPLICATION_ID"].ToString();
                display_NameTextBox.Text = dataRow["DISPLAY_NAME"].ToString();
                application_NameTextBox.Text = dataRow["APPLICATION_NAME"].ToString();
                application_URLTextBox.Text = dataRow["APPLICATION_URL"].ToString();
                login_URLTextBox.Text = dataRow["LOGIN_URL"].ToString();
                usernameTextBox.Text = dataRow["USERNAME"].ToString();
                emailTextBox.Text = dataRow["EMAIL"].ToString();
                passwordTextBox.Text = dataRow["PASSWORD"].ToString();
                software_KeyTextBox.Text = dataRow["SOFTWARE_KEY"].ToString();
                cD_KeyTextBox.Text = dataRow["CD_KEY"].ToString();
                general_Key_1TextBox.Text = dataRow["GENERAL_KEY_1"].ToString();
                general_Key_2TextBox.Text = dataRow["GENERAL_KEY_2"].ToString();
                general_Key_3TextBox.Text = dataRow["GENERAL_KEY_3"].ToString();
                general_Key_4TextBox.Text = dataRow["GENERAL_KEY_4"].ToString();
                general_Key_5TextBox.Text = dataRow["GENERAL_KEY_5"].ToString();
                has_Credit_Card_InfoCheckBox.Checked = (bool)dataRow["HAS_CREDIT_CARD_INFO"];
                has_Address_InfoCheckBox.Checked = (bool)dataRow["HAS_ADDRESS_INFO"];
                has_Phone_NumberCheckBox.Checked = (bool)dataRow["HAS_PHONE_NUMBER"];
                is_Auto_RenewalCheckBox.Checked = (bool)dataRow["IS_AUTO_RENEWAL"];
                notesTextBox.Text = dataRow["NOTES"].ToString();
            }
            catch (Exception e)
            {
                MessageBox.Show("Failed to populate form with record data. Error = " + e.Message);
                ClearEntryFields();
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrWhiteSpace(application_IDTextBox.Text))
            {
                DialogResult result = MessageBox.Show("Are you sure you want to delete?", "Caption", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.ServiceNotification);
                if (result == DialogResult.Yes)
                {
                    int index = FindRecordByApplicationIDAndReturnIndexInDataSet(application_IDTextBox.Text);
                    if (index >= 0)
                    {
                        forgetMeNotDataSet.ForgetMeNot.Rows[index].Delete();
                        this.forgetMeNotBindingSource.EndEdit();
                        this.tableAdapterManager.UpdateAll(this.forgetMeNotDataSet);
                        ClearEntryFields();
                    }
                }
            }
            else
            {
                MessageBox.Show("No Record Loaded In Editor");
            }
        }

        private void forgetMeNotBindingNavigatorSaveItem_Click_1(object sender, EventArgs e)
        {
            this.Validate();
            this.forgetMeNotBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.forgetMeNotDataSet);

        }

        private void button4_Click(object sender, EventArgs e)
        {
            ExportToExcel(forgetMeNotDataGridView);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ClearEntryFields();
        }

    }
}
