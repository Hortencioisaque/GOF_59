using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Microsoft.VisualBasic;

namespace G
{
    public partial class GOF_59 : Form
    {
        string strID = string.Empty;//Employees - ID from the Employees table

        string mdfFile = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\horte\source\repos\GOF_59\GDB.mdb";
        //string mdfFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\GDB.mdb";//path of the pc/user Documents folder


        //***************************************************************************************************************************************************************************************************************************
        //
        //
        //
        //START On Load -  showing Expenses category, Bill type & Company name when the program starts
        //
        //
        //
        //***************************************************************************************************************************************************************************************************************************
        private void Form1_Load(object sender, EventArgs e)// On load, displaying 
        {
            //
            //START of Employees On Load -  ComboBox to show Employees
            //
            //
            //START of Updating ComboBox to show Employees
            //
            List<string> ListOfNames = new List<string>();
            //
            //select command
            //
            //string mdfFile1 = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\IsaqueH\Documents\GDB.mdb";


            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
            {
                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Employees", connection))
                {
                    connection.Open();

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        object ComboNameValue = row["Combo_Names"];
                        string strName = ComboNameValue + "";

                        if (!ListOfNames.Contains(strName))
                        {
                            comboBoxEmployees.Items.Add(strName);//adding items to the comboBox Employees Tab
                            ComboBoxWage.Items.Add(strName);//adding items to the comboBox in the *WAGE TAB

                            comboBoxEmployees.AutoCompleteMode = AutoCompleteMode.Suggest;//ComboBox Auto suggestion when typing first characters 
                            comboBoxEmployees.AutoCompleteSource = AutoCompleteSource.ListItems;//ComboBox Auto suggestion when typing first characters 
                            ComboBoxWage.AutoCompleteMode = AutoCompleteMode.Suggest;//ComboBox Auto suggestion when typing first characters 
                            ComboBoxWage.AutoCompleteSource = AutoCompleteSource.ListItems;//ComboBox Auto suggestion when typing first characters 


                            comboBoxEmployees.Sorted = true;
                            ComboBoxWage.Sorted = true;
                            ListOfNames.Add(strName);
                        }
                    }
                }
            }

            //
            //end of select  command
            //
            //

            //
            //END of Updating ComboBox to show Employees
            //


            //
            //END of Employees On load -  ComboBox to show Employees
            //


            //
            //START of Expenses category combo box On Load -  ComboBox to show Expenses category, Bill type & Company name from the database
            //
            //***********************************************************************************************************************************************************************************************************************

            //
            //START of Updating ComboBox to show Expenses category
            //
            //
            //select command
            //



            comboBoxExpensesCategory.AutoCompleteMode = AutoCompleteMode.Suggest;//ComboBox Auto suggestion when typing first characters 
            comboBoxExpensesCategory.AutoCompleteSource = AutoCompleteSource.ListItems;//ComboBox Auto suggestion when typing first characters 

            comboBoxbuttonExpensesBill_Type.AutoCompleteMode = AutoCompleteMode.Suggest;//ComboBox Auto suggestion when typing first characters 
            comboBoxbuttonExpensesBill_Type.AutoCompleteSource = AutoCompleteSource.ListItems;//ComboBox Auto suggestion when typing first characters 

            comboBoxExpensesCompanies.AutoCompleteMode = AutoCompleteMode.Suggest;//ComboBox Auto suggestion when typing first characters 
            comboBoxExpensesCompanies.AutoCompleteSource = AutoCompleteSource.ListItems;//ComboBox Auto suggestion when typing first characters 

            comboBoxExpensesCategory.Sorted = true;

            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
            {

                //
                //collecting from expenses
                //
                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Expenses_Category", connection))
                {
                    connection.Open();

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        object CategoryNameValue = row["Category_EXPENSES_Name"];
                        string strExpensesCategoryName = CategoryNameValue + "";//

                        if (!ListOfNames.Contains(strExpensesCategoryName))
                        {
                            comboBoxExpensesCategory.Items.Add(strExpensesCategoryName);//adding items to the comboBox Expenses category Tab
                            ListOfNames.Add(strExpensesCategoryName);
                        }

                    }
                }
                //
                //End of collecting from expenses Category
                //

                //
                //collecting from expenses Bill Type
                //
                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Expenses_Bill_type", connection))
                {

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        object BillTypeNameValue = row["Bill_type_EXPENSES_Name"];
                        string strExpensesBillTypeName = BillTypeNameValue + "";//

                        if (!ListOfNames.Contains(strExpensesBillTypeName))
                        {
                            comboBoxbuttonExpensesBill_Type.Items.Add(strExpensesBillTypeName);//adding items to the comboBox Expenses category Tab
                            ListOfNames.Add(strExpensesBillTypeName);
                        }
                    }
                }
                //
                //End of collecting from expenses Bill Type
                //

                //
                //collecting from expenses Bill Type
                //
                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Expenses_Companies", connection))
                {

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        object CategoryNameValue = row["Company_EXPENSES_Name"];
                        string strExpensesCompanyName = CategoryNameValue + "";//

                        if (!ListOfNames.Contains(strExpensesCompanyName))
                        {
                            comboBoxExpensesCompanies.Items.Add(strExpensesCompanyName);//adding items to the comboBox Expenses category Tab
                            ListOfNames.Add(strExpensesCompanyName);
                        }
                    }
                }
                //
                //End of collecting from expenses Bill Type
                //

            }

            //
            //end of select  command
            //
            //

            //
            //END of Updating ComboBox to show  Expenses category
            //
            //END of Expenses category combo box On Load -ComboBox to show Expenses category, Bill type & Company name
            //

        }

        //***************************************************************************************************************************************************************************************************************************
        //
        //
        //
        //END of On Load -  showing Expenses category, Bill type & Company name when the program starts
        //
        //
        //
        //***************************************************************************************************************************************************************************************************************************


        public GOF_59()
        {
            InitializeComponent();
        }
        //***************************************************************************************************************************************************************************************************************************
        //
        //
        //
        //START of Employees Below  
        //
        //
        //
        //***************************************************************************************************************************************************************************************************************************


        private void button5_Click_1(object sender, EventArgs e)//Employees // insert and select
        {

            //
            //
            //START of Adding a new Employee
            //
            //

            richTextBoxEmp.Clear();

            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider =Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))//using connection
            {
                using (OleDbCommand insertCommand = new OleDbCommand("INSERT INTO Employees ([First_Name],[Middle_Names],[Last_Name],[Address],[Address2],[City],[Country],[Post_Code],[NINO],[Email],[Start_Date],[Combo_Names]) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", connection))//insert command
                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Employees", connection))
                {
                    connection.Open();

                    //
                    //start of checking if entry exists in database
                    //
                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    int iEntryCount = 0;

                    foreach (DataRow row in table.Rows)
                    {

                        object nameValue = row["First_Name"];
                        object Middle_names = row["Middle_Names"];
                        object Last_Name = row["Last_Name"];
                        object Nino = row["NINO"];

                        string strName = nameValue.ToString();
                        string strM_Name = Middle_names.ToString();
                        string strL_Name = Last_Name.ToString();
                        string strNino = Nino.ToString();

                        //
                        //End of checking if entry exists in database
                        //

                        //
                        //inserting into the database
                        //
                        if (strName.Contains(Emp1.Text) && strM_Name.Contains(Emp2.Text) && strL_Name.Contains(Emp3.Text) && strNino.Contains(Emp9.Text))
                        {
                            MessageBox.Show("Employee already exist");

                        }
                        else
                        {
                            iEntryCount++;// To stop the loop and save only 1 time the new entry

                            if (iEntryCount == 1)
                            {
                                insertCommand.Parameters.AddWithValue("@First_Name", Emp1.Text);
                                insertCommand.Parameters.AddWithValue("@Middle_Names", Emp2.Text);
                                insertCommand.Parameters.AddWithValue("@Last_Name", Emp3.Text);
                                insertCommand.Parameters.AddWithValue("@Address", Emp4.Text);
                                insertCommand.Parameters.AddWithValue("@Address2", Emp5.Text);
                                insertCommand.Parameters.AddWithValue("@City", Emp6.Text);
                                insertCommand.Parameters.AddWithValue("@Country", Emp7.Text);
                                insertCommand.Parameters.AddWithValue("@Post_Code", Emp8.Text);
                                insertCommand.Parameters.AddWithValue("@NINO", Emp9.Text);
                                insertCommand.Parameters.AddWithValue("@Email", Emp10.Text);
                                insertCommand.Parameters.AddWithValue("@Start_Date", Emp11.Text);
                                insertCommand.Parameters.AddWithValue("Combo_Names", Emp1.Text + " " + Emp2.Text + " " + Emp3.Text);

                                insertCommand.ExecuteNonQuery();

                                string strComboNames = Emp1.Text + " " + Emp2.Text + " " + Emp3.Text;

                                //
                                //end of inserting into database
                                //

                                //
                                //START of Showing what has been added to the database
                                //
                                richTextBoxEmp.AppendText("Details of the New Record Added :    " + Environment.NewLine);
                                richTextBoxEmp.AppendText(Environment.NewLine);
                                richTextBoxEmp.AppendText(Emp1.Text + "    " + Environment.NewLine);
                                richTextBoxEmp.AppendText(Emp2.Text + "    " + Environment.NewLine);
                                richTextBoxEmp.AppendText(Emp3.Text + "    " + Environment.NewLine);
                                richTextBoxEmp.AppendText(Emp4.Text + "    " + Environment.NewLine);
                                richTextBoxEmp.AppendText(Emp5.Text + "    " + Environment.NewLine);
                                richTextBoxEmp.AppendText(Emp6.Text + "    " + Environment.NewLine);
                                richTextBoxEmp.AppendText(Emp7.Text + "    " + Environment.NewLine);
                                richTextBoxEmp.AppendText(Emp8.Text + "    " + Environment.NewLine);
                                richTextBoxEmp.AppendText(Emp9.Text + "    " + Environment.NewLine);
                                richTextBoxEmp.AppendText(Emp10.Text + "    " + Environment.NewLine);
                                richTextBoxEmp.AppendText(Emp11.Text + "    " + Environment.NewLine);
                                //
                                //END of Showing what has been added to the database
                                //

                                //
                                //adding to the comboBox
                                //
                                comboBoxEmployees.Items.Add(strComboNames);
                                ComboBoxWage.Items.Add(strComboNames);//adding items to the comboBox WAGE TAB as the employers table are updated
                                //
                                //end of adding to the comboBox
                                //
                            }
                        }
                    }
                }
            }
            //
            //END of Adding a new Employee
            //

            //
            //START of clearing fields and comboBox to update comboBox with new records
            //
            comboBoxEmployees.Items.Clear();
            comboBoxEmployees.ResetText();
            ComboBoxWage.Items.Clear();
            ComboBoxWage.ResetText();
            Emp1.Clear();
            Emp2.Clear();
            Emp3.Clear();
            Emp4.Clear();
            Emp5.Clear();
            Emp6.Clear();
            Emp7.Clear();
            Emp8.Clear();
            Emp9.Clear();
            Emp10.Clear();
            Emp11.Clear();
            EmpID.Clear();
            //
            //END of clearing fields and comboBox to update comboBox with new records
            //

            //
            //START of Updating ComboBox to show Employees
            //
            List<string> ListOfNames = new List<string>();
            //
            //select command
            //
            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
            {
                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Employees", connection))
                {
                    connection.Open();

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        object ID = row["ID"];
                        object nameValue = row["First_Name"];
                        object Middle_names = row["Middle_Names"];
                        object Last_Name = row["Last_Name"];
                        strID = ID + "";
                        string strName = nameValue + " " + Middle_names + " " + Last_Name;//

                        if (!ListOfNames.Contains(strName))
                        {
                            comboBoxEmployees.Items.Add(strName);//adding items to the combobox
                            ComboBoxWage.Items.Add(strName);//adding items to the comboBox WAGE TAB
                            ListOfNames.Add(strName);
                        }
                    }
                }
            }
            //
            //end of selcet  command
            //
            //

            //
            //END of Updating ComboBox to show Employees
            //
        }
        //
        //
        //Start of Editing Employees by selecting employee in the ComboBox and displaying in fields for editing.
        //
        //
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //clearing text boxes when jumping from on employee to another.
            richTextBoxEmp.Clear();
            Emp1.Clear();
            Emp2.Clear();
            Emp3.Clear();
            Emp4.Clear();
            Emp5.Clear();
            Emp6.Clear();
            Emp7.Clear();
            Emp8.Clear();
            Emp9.Clear();
            Emp10.Clear();
            Emp11.Clear();
            EmpID.Clear();

            if (!String.IsNullOrEmpty(comboBoxEmployees.Text))
            {
                //
                //select command
                //
                using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
                {
                    using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Employees", connection))
                    {
                        connection.Open();

                        DataTable table = new DataTable();
                        OleDbDataAdapter adapter = new OleDbDataAdapter();
                        adapter.SelectCommand = selectCommand;
                        adapter.Fill(table);


                        string strComboBox = comboBoxEmployees.Text;// string inside combo box... employee selected

                        string strComboNames = strComboBox;//.Substring(0, strComboBox.IndexOf(" "));// collecting only the employee ID

                        //richTextBoxEmp.AppendText(strComboBox);

                        foreach (DataRow row in table.Rows)
                        {

                            //
                            //Collecting all from employees
                            //
                            object ID = row["ID"];
                            object firstName = row["First_Name"];
                            object MiddleName = row["Middle_Names"];
                            object LastName = row["Last_Name"];
                            object Address = row["Address"];
                            object Address2 = row["Address2"];
                            object City = row["City"];
                            object Country = row["Country"];
                            object PostCode = row["Post_Code"];
                            object Nino = row["NINO"];
                            object Email = row["Email"];
                            object StartDate = row["Start_Date"];
                            object ComboNames = row["Combo_Names"];
                            //
                            //Collected all from employees
                            //
                            //
                            //START of inserting selected employee into fields for editing
                            //
                            if (ComboNames.ToString() == strComboNames)
                            {
                                EmpID.AppendText(ID.ToString());
                                Emp1.AppendText(firstName.ToString());
                                Emp2.AppendText(MiddleName.ToString());
                                Emp3.AppendText(LastName.ToString());
                                Emp4.AppendText(Address.ToString());
                                Emp5.AppendText(Address2.ToString());
                                Emp6.AppendText(City.ToString());
                                Emp7.AppendText(Country.ToString());
                                Emp8.AppendText(PostCode.ToString());
                                Emp9.AppendText(Nino.ToString());
                                Emp10.AppendText(Email.ToString());
                                Emp11.AppendText(StartDate.ToString());

                            }

                            //
                            //END of inserting selected employee into fields for editing
                            //

                        }
                    }
                }
            }
        }

        //
        //End of Editing Employees by selecting employee in the ComboBox and displaying in fields for editing.
        //

        //
        //Start of update employees DB table fields
        //
        private void button4_Click(object sender, EventArgs e)//update - employees table
        {
            richTextBoxEmp.Clear();
            //
            //Update command
            //
            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
            {
                using (OleDbCommand updateCommand = new OleDbCommand("UPDATE Employees SET [First_Name] = ?, [Middle_Names] = ?, [Last_Name] = ?, [Address] = ?, [Address2] = ?, [City] = ?, [Country] = ?, [Post_Code] = ?, [NINO] = ?, [Email] = ?, [Start_Date] = ?, [Combo_Names] = ? WHERE [ID] = ?", connection))
                {
                    connection.Open();

                    //updateCommand.Parameters.AddWithValue("@Category", "Drink");

                    updateCommand.Parameters.AddWithValue("@First_Name", Emp1.Text);
                    updateCommand.Parameters.AddWithValue("@Middle_Names", Emp2.Text);
                    updateCommand.Parameters.AddWithValue("@Last_Name", Emp3.Text);
                    updateCommand.Parameters.AddWithValue("@Address", Emp4.Text);
                    updateCommand.Parameters.AddWithValue("@Address", Emp5.Text);
                    updateCommand.Parameters.AddWithValue("@City", Emp6.Text);
                    updateCommand.Parameters.AddWithValue("@City", Emp7.Text);
                    updateCommand.Parameters.AddWithValue("@Post_Code", Emp8.Text);
                    updateCommand.Parameters.AddWithValue("@NINO", Emp9.Text);
                    updateCommand.Parameters.AddWithValue("@Email", Emp10.Text);
                    updateCommand.Parameters.AddWithValue("@Start_Date", Emp11.Text);

                    string strComboNames = Emp1.Text + " " + Emp2.Text + " " + Emp3.Text;
                    updateCommand.Parameters.AddWithValue("@Combo_Names", strComboNames);
                    updateCommand.Parameters.AddWithValue("@ID", EmpID.Text);

                    updateCommand.ExecuteNonQuery();

                    //
                    //START of Showing what has been updated into the database
                    //
                    richTextBoxEmp.AppendText("Records Updated:    " + Environment.NewLine);
                    richTextBoxEmp.AppendText(Environment.NewLine);
                    richTextBoxEmp.AppendText(Emp1.Text + "    " + Environment.NewLine);
                    richTextBoxEmp.AppendText(Emp2.Text + "    " + Environment.NewLine);
                    richTextBoxEmp.AppendText(Emp3.Text + "    " + Environment.NewLine);
                    richTextBoxEmp.AppendText(Emp4.Text + "    " + Environment.NewLine);
                    richTextBoxEmp.AppendText(Emp5.Text + "    " + Environment.NewLine);
                    richTextBoxEmp.AppendText(Emp6.Text + "    " + Environment.NewLine);
                    richTextBoxEmp.AppendText(Emp7.Text + "    " + Environment.NewLine);
                    richTextBoxEmp.AppendText(Emp8.Text + "    " + Environment.NewLine);
                    richTextBoxEmp.AppendText(Emp9.Text + "    " + Environment.NewLine);
                    richTextBoxEmp.AppendText(Emp10.Text + "    " + Environment.NewLine);
                    richTextBoxEmp.AppendText(Emp11.Text + "    " + Environment.NewLine);
                    richTextBoxEmp.AppendText(strComboNames + "    " + Environment.NewLine);

                    //
                    //END of Showing what has been updated into the database
                    //
                }
            }
            //
            //end of update cmd
            //

            //
            //START of clearing fields and comboBox to update comboBox with new records
            //
            comboBoxEmployees.Items.Clear();
            comboBoxEmployees.ResetText();
            ComboBoxWage.Items.Clear();
            ComboBoxWage.ResetText();
            Emp1.Clear();
            Emp2.Clear();
            Emp3.Clear();
            Emp4.Clear();
            Emp5.Clear();
            Emp6.Clear();
            Emp7.Clear();
            Emp8.Clear();
            Emp9.Clear();
            Emp10.Clear();
            Emp11.Clear();
            EmpID.Clear();
            //
            //END of clearing fields and comboBox to update comboBox with new records
            //

            //
            //START of Updating ComboBox to show Employees
            //
            List<string> ListOfNames = new List<string>();
            //
            //select command
            //
            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
            {
                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Employees", connection))
                {
                    connection.Open();

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        object ID = row["ID"];
                        object nameValue = row["First_Name"];
                        object Middle_names = row["Middle_Names"];
                        object Last_Name = row["Last_Name"];
                        strID = ID + "";
                        string strName = nameValue + " " + Middle_names + " " + Last_Name;//

                        if (!ListOfNames.Contains(strName))
                        {
                            comboBoxEmployees.Items.Add(strName);//adding items to the combobox
                            ComboBoxWage.Items.Add(strName);//adding items to the comboBox WAGE TAB
                            ListOfNames.Add(strName);
                        }
                    }
                }
            }
            //
            //end of selcet  command
            //
            //

            //
            //END of Updating ComboBox to show Employees
            //
        }
        //
        //End of update employees table
        //

        //
        //start of deleting emplyee records from databsae
        //
        private void button3_Click(object sender, EventArgs e)//Delete - emplyee records from databsae
        {
            richTextBoxEmp.Clear();
            //
            //START of DELETING ymployee record from the database.
            //            
            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
            {
                using (OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM Employees WHERE [ID] = ?", connection))
                {
                    connection.Open();

                    deleteCommand.Parameters.AddWithValue("@ID", EmpID.Text);

                    deleteCommand.ExecuteNonQuery();
                }
            }

            //
            //START of Showing what has been updated into the database
            //
            richTextBoxEmp.AppendText("Records Deleted Successfully:    " + Environment.NewLine);
            richTextBoxEmp.AppendText(Environment.NewLine);
            richTextBoxEmp.AppendText(Emp1.Text + "    " + Environment.NewLine);
            richTextBoxEmp.AppendText(Emp2.Text + "    " + Environment.NewLine);
            richTextBoxEmp.AppendText(Emp3.Text + "    " + Environment.NewLine);
            richTextBoxEmp.AppendText(Emp4.Text + "    " + Environment.NewLine);
            richTextBoxEmp.AppendText(Emp5.Text + "    " + Environment.NewLine);
            richTextBoxEmp.AppendText(Emp6.Text + "    " + Environment.NewLine);
            richTextBoxEmp.AppendText(Emp7.Text + "    " + Environment.NewLine);
            richTextBoxEmp.AppendText(Emp8.Text + "    " + Environment.NewLine);
            richTextBoxEmp.AppendText(Emp9.Text + "    " + Environment.NewLine);
            richTextBoxEmp.AppendText(Emp10.Text + "    " + Environment.NewLine);
            richTextBoxEmp.AppendText(Emp11.Text + "    " + Environment.NewLine);
            //
            //END of Showing what has been updated into the database
            //

            //
            //END of DELETING ymployee record from the database.
            //

            //
            //START of clearing fields and comboBox to update comboBox with new records
            //
            comboBoxEmployees.Items.Clear();
            comboBoxEmployees.ResetText();
            ComboBoxWage.Items.Clear();
            ComboBoxWage.ResetText();
            Emp1.Clear();
            Emp2.Clear();
            Emp3.Clear();
            Emp4.Clear();
            Emp5.Clear();
            Emp6.Clear();
            Emp7.Clear();
            Emp8.Clear();
            Emp9.Clear();
            Emp10.Clear();
            Emp11.Clear();
            EmpID.Clear();
            //
            //END of clearing fields and comboBox to update comboBox with new records
            //

            //
            //START of Updating ComboBox to show Employees
            //
            List<string> ListOfNames = new List<string>();
            //
            //select command
            //

            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
            {
                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Employees", connection))
                {
                    connection.Open();

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        object ID = row["ID"];
                        object nameValue = row["First_Name"];
                        object Middle_names = row["Middle_Names"];
                        object Last_Name = row["Last_Name"];
                        strID = ID + "";
                        string strName = nameValue + " " + Middle_names + " " + Last_Name;//

                        if (!ListOfNames.Contains(strName))
                        {
                            comboBoxEmployees.Items.Add(strName);//adding items to the combobox
                            ComboBoxWage.Items.Add(strName);//adding items to the comboBox WAGE TAB
                            ListOfNames.Add(strName);
                        }
                    }
                }
            }
            //
            //end of selcet  command
            //
            //

            //
            //END of Updating ComboBox to show Employees
            //
        }
        //
        //END of deleting emplyee records from databsae
        //

        //***************************************************************************************************************************************************************************************************************************
        //
        //
        //
        //END of Employees along with combo box for the wage to update as employees are updated
        //
        //
        //
        //***************************************************************************************************************************************************************************************************************************


        // ################################################################   END OF Employees   #################################################################################



        //***************************************************************************************************************************************************************************************************************************
        //
        //
        //
        //
        //START of EXPENSES
        //
        //
        //
        //
        //***************************************************************************************************************************************************************************************************************************



        private void button14_Click(object sender, EventArgs e)//Insert/Select A new record to the Expenses_Category
        {
            //
            //Adding a new category Expenses to the database using a input message box
            //
            List<string> List_Category_EXPENSES_Name = new List<string>();// to add all the categories to a list for latter to be checked if new entry already exists in the table
            string NewCategoryExpensesValue = Interaction.InputBox("Please insert a new category", "New Category", "");

            richTextBox3.AppendText(NewCategoryExpensesValue);

            //
            //End of collecting a new category Expenses to add in the database using a inputBox message box
            //
            if (NewCategoryExpensesValue != "" && comboBoxExpensesCategory.Text == "")
            {

                richTextBox3.Clear();

                //
                //START of new records to the Expenses_Category Table.
                //
                using (OleDbConnection connection = new OleDbConnection(string.Format("Provider =Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))//using connection
                {

                    using (OleDbCommand insertCommand = new OleDbCommand("INSERT INTO Expenses_Category ([Category_EXPENSES_Name]) VALUES (?)", connection))//insert command
                    using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Expenses_Category", connection))
                    {

                        connection.Open();

                        //
                        //start of checking if entry exists in database
                        //
                        DataTable table = new DataTable();
                        OleDbDataAdapter adapter = new OleDbDataAdapter();
                        adapter.SelectCommand = selectCommand;
                        adapter.Fill(table);

                        foreach (DataRow row in table.Rows)
                        {

                            object nameValue = row["Category_EXPENSES_Name"];

                            string strName = nameValue.ToString();

                            if (!List_Category_EXPENSES_Name.Contains(strName))
                            {
                                List_Category_EXPENSES_Name.Add(strName);
                            }

                            //
                            //End of checking if entry exists in database
                            //
                        }

                        if (List_Category_EXPENSES_Name.Contains(NewCategoryExpensesValue))
                        {
                            MessageBox.Show("Category already exist");
                        }
                        //
                        //inserting into the database
                        //
                        else
                        {
                            insertCommand.Parameters.AddWithValue("Category_EXPENSES_Name", NewCategoryExpensesValue);

                            insertCommand.ExecuteNonQuery();

                            connection.Close();

                            //
                            //end of inserting into database
                            //

                            //
                            //START of Showing what has been added to the database
                            //
                            richTextBox3.AppendText("Records Inserted:    " + Environment.NewLine);
                            richTextBox3.AppendText(Environment.NewLine);
                            richTextBox3.AppendText(NewCategoryExpensesValue + "    " + Environment.NewLine);
                            //
                            //END of Showing what has been added to the database
                            //

                            //
                            //adding to the comboBox
                            //

                            comboBoxExpensesCategory.Items.Add(NewCategoryExpensesValue);

                            //
                            //end of adding to the comboBox
                            //
                        }



                    }
                }
                //
                //END of Adding a new Expenses category
                //

                //
                //START of clearing fields and comboBox to update comboBox with new records
                //
                /*
                comboBoxExpensesCategory.Items.Clear();
                comboBoxExpensesCategory.ResetText();
                */
                //
                //END of clearing fields and comboBox to update comboBox with new records
                //

                // END of on click to check for an existing entri and add a new entry to categories expenses

            }
        }
        //
        //end of Add a new expenses category
        //


        //
        //Delete expenses category
        //

        private void button30_Click(object sender, EventArgs e)//Delete - expenses category
        {
            richTextBox3.Clear();
            //
            //START of DELETING Expenses_Category  record from the database.
            //            

            if (comboBoxExpensesCategory.Text != "")
            {
                using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
                {
                    using (OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM Expenses_Category WHERE [Category_EXPENSES_Name] = ?", connection))
                    {
                        connection.Open();

                        deleteCommand.Parameters.AddWithValue("@Category_EXPENSES_Name", comboBoxExpensesCategory.Text);

                        deleteCommand.ExecuteNonQuery();
                    }
                }

                //
                //START of Showing what has been updated into the database
                //
                richTextBox3.AppendText("Records Deleted Successfully:    " + Environment.NewLine);
                richTextBox3.AppendText(comboBoxExpensesCategory.Text + Environment.NewLine);
                comboBoxExpensesCategory.Items.Remove(comboBoxExpensesCategory.Text);


                //
                //END of Showing what has been updated into the database
                //

                //
                //END of DELETING Expenses_Category record from the database.
                //

                //
                //START of clearing fields and comboBox to update comboBox with new records
                //
                /*
                comboBoxEmployees.Items.Clear();
                comboBoxEmployees.ResetText();
                */
                //
                //END of clearing fields and comboBox to update comboBox with new records
                //
            }
            else
            {
                MessageBox.Show("Please select an entry");
            }
        }

        //
        //END Delete expenses category
        //


        //
        //ADD a new expenses Bill Type
        //

        private void buttonAddExpensesBill_Type_Click(object sender, EventArgs e) // ADD a new expenses Bill Type
        {

            //
            //Adding a new category Expenses to the database using a input message box
            //
            List<string> List_Bill_Type_EXPENSES_Name = new List<string>();// to add all the categories to a list for latter to be checked if new entry already exists in the table
            string NewBillTypeExpensesValue = Interaction.InputBox("Please insert a new category", "New Ctegory", "");

            richTextBox3.AppendText(NewBillTypeExpensesValue);

            //
            //End of collecting a new category Expenses to add in the database using a inputBox message box
            //
            if (NewBillTypeExpensesValue != "" && comboBoxbuttonExpensesBill_Type.Text == "")
            {

                richTextBox3.Clear();

                //
                //START of new records to the Expenses_Category Table.
                //
                using (OleDbConnection connection = new OleDbConnection(string.Format("Provider =Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))//using connection
                {

                    using (OleDbCommand insertCommand = new OleDbCommand("INSERT INTO Expenses_Bill_type ([Bill_type_EXPENSES_Name]) VALUES (?)", connection))//insert command
                    using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Expenses_Bill_type", connection))
                    {

                        connection.Open();

                        //
                        //start of checking if entry exists in database
                        //
                        DataTable table = new DataTable();
                        OleDbDataAdapter adapter = new OleDbDataAdapter();
                        adapter.SelectCommand = selectCommand;
                        adapter.Fill(table);

                        foreach (DataRow row in table.Rows)
                        {

                            object nameValue = row["Bill_type_EXPENSES_Name"];

                            string strName = nameValue.ToString();

                            if (!List_Bill_Type_EXPENSES_Name.Contains(strName))
                            {
                                List_Bill_Type_EXPENSES_Name.Add(strName);
                            }

                            //
                            //End of checking if entry exists in database
                            //
                        }

                        if (List_Bill_Type_EXPENSES_Name.Contains(NewBillTypeExpensesValue))
                        {
                            MessageBox.Show("Category already exist");
                        }
                        //
                        //inserting into the database
                        //
                        else
                        {
                            insertCommand.Parameters.AddWithValue("Bill_type_EXPENSES_Name", NewBillTypeExpensesValue);

                            insertCommand.ExecuteNonQuery();

                            connection.Close();

                            //
                            //end of inserting into database
                            //

                            //
                            //START of Showing what has been added to the database
                            //
                            richTextBox3.AppendText("Records Inserted:    " + Environment.NewLine);
                            richTextBox3.AppendText(Environment.NewLine);
                            richTextBox3.AppendText(NewBillTypeExpensesValue + "    " + Environment.NewLine);
                            //
                            //END of Showing what has been added to the database
                            //

                            //
                            //adding to the comboBox
                            //

                            comboBoxbuttonExpensesBill_Type.Items.Add(NewBillTypeExpensesValue);

                            //
                            //end of adding to the comboBox
                            //
                        }



                    }
                }
            }
        }

        //
        //END of ADD a new expenses Bill Type
        //


        //
        //Delete expenses Bill Type
        //

        private void button31_Click(object sender, EventArgs e) //Delete expenses Bill Type
        {
            richTextBox3.Clear();
            //
            //START of DELETING Expenses_Category  record from the database.
            //            

            if (comboBoxbuttonExpensesBill_Type.Text != "")
            {
                using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
                {
                    using (OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM Expenses_Bill_type WHERE [Bill_type_EXPENSES_Name] = ?", connection))
                    {
                        connection.Open();

                        deleteCommand.Parameters.AddWithValue("@Bill_type_EXPENSES_Name", comboBoxbuttonExpensesBill_Type.Text);

                        deleteCommand.ExecuteNonQuery();
                    }
                }

                //
                //START of Showing what has been updated into the database
                //
                richTextBox3.AppendText("Records Deleted Successfully:    " + Environment.NewLine);
                richTextBox3.AppendText(comboBoxbuttonExpensesBill_Type.Text + Environment.NewLine);
                comboBoxbuttonExpensesBill_Type.Items.Remove(comboBoxbuttonExpensesBill_Type.Text);


                //
                //END of Showing what has been updated into the database
                //

                //
                //END of DELETING Expenses_Category record from the database.
                //

                //
                //START of clearing fields and comboBox to update comboBox with new records
                //
                /*
                comboBoxEmployees.Items.Clear();
                comboBoxEmployees.ResetText();
                */
                //
                //END of clearing fields and comboBox to update comboBox with new records
                //
            }
            else
            {
                MessageBox.Show("Please select an entry");
            }
        }

        //
        //END Delete expenses Bill Type
        //

        //
        //ADD a new expenses Campany name
        //

        private void button15_Click(object sender, EventArgs e) //ADD a new expenses Campany name
        {
            //
            //Adding a new category Expenses to the database using a input message box
            //
            List<string> List_Company_EXPENSES_Name = new List<string>();// to add all the categories to a list for latter to be checked if new entry already exists in the table
            string NewComanyNameExpensesValue = Interaction.InputBox("Please insert a new category", "New Category", "");

            richTextBox3.AppendText(NewComanyNameExpensesValue);

            //
            //End of collecting a new category Expenses to add in the database using a inputBox message box
            //
            if (NewComanyNameExpensesValue != "" && comboBoxExpensesCompanies.Text == "")
            {

                richTextBox3.Clear();

                //
                //START of new records to the Expenses_Companies Table.
                //
                using (OleDbConnection connection = new OleDbConnection(string.Format("Provider =Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))//using connection
                {

                    using (OleDbCommand insertCommand = new OleDbCommand("INSERT INTO Expenses_Companies ([Company_EXPENSES_Name]) VALUES (?)", connection))//insert command
                    using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Expenses_Companies", connection))
                    {

                        connection.Open();

                        //
                        //start of checking if entry exists in database
                        //
                        DataTable table = new DataTable();
                        OleDbDataAdapter adapter = new OleDbDataAdapter();
                        adapter.SelectCommand = selectCommand;
                        adapter.Fill(table);

                        foreach (DataRow row in table.Rows)
                        {

                            object nameValue = row["Company_EXPENSES_Name"];

                            string strName = nameValue.ToString();

                            if (!List_Company_EXPENSES_Name.Contains(strName))
                            {
                                List_Company_EXPENSES_Name.Add(strName);
                            }

                            //
                            //End of checking if entry exists in database
                            //
                        }

                        if (List_Company_EXPENSES_Name.Contains(NewComanyNameExpensesValue))
                        {
                            MessageBox.Show("Category already exist");
                        }
                        //
                        //inserting into the database
                        //
                        else
                        {
                            insertCommand.Parameters.AddWithValue("Company_EXPENSES_Name", NewComanyNameExpensesValue);

                            insertCommand.ExecuteNonQuery();

                            connection.Close();

                            //
                            //end of inserting into database
                            //

                            //
                            //START of Showing what has been added to the database
                            //
                            richTextBox3.AppendText("Records Inserted:    " + Environment.NewLine);
                            richTextBox3.AppendText(Environment.NewLine);
                            richTextBox3.AppendText(NewComanyNameExpensesValue + "    " + Environment.NewLine);
                            //
                            //END of Showing what has been added to the database
                            //

                            //
                            //adding to the comboBox
                            //

                            comboBoxExpensesCompanies.Items.Add(NewComanyNameExpensesValue);

                            //
                            //end of adding to the comboBox
                            //
                        }



                    }
                }
                //
                //END of Adding a new Expenses category
                //

                //
                //START of clearing fields and comboBox to update comboBox with new records
                //
                /*
                comboBoxExpensesCompanies.Items.Clear();
                comboBoxExpensesCompanies.ResetText();
                */
                //
                //END of clearing fields and comboBox to update comboBox with new records
                //

                // END of on click to check for an existing entri and add a new entry to categories expenses

            }
        }

        //
        //ADD a new expenses Campany name
        //


        //
        //Delete expenses Capany name
        //

        private void button32_Click(object sender, EventArgs e) //Delete expenses Capany name
        {
            richTextBox3.Clear();
            //
            //START of DELETING Expenses_Category  record from the database.
            //            

            if (comboBoxExpensesCompanies.Text != "")
            {
                using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
                {
                    using (OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM Expenses_Companies WHERE [Company_EXPENSES_Name] = ?", connection))
                    {
                        connection.Open();

                        deleteCommand.Parameters.AddWithValue("@Company_EXPENSES_Name", comboBoxExpensesCompanies.Text);

                        deleteCommand.ExecuteNonQuery();
                    }
                }

                //
                //START of Showing what has been updated into the database
                //
                richTextBox3.AppendText("Records Deleted Successfully:    " + Environment.NewLine);
                richTextBox3.AppendText(comboBoxExpensesCompanies.Text + Environment.NewLine);
                comboBoxExpensesCompanies.Items.Remove(comboBoxExpensesCompanies.Text);


                //
                //END of Showing what has been updated into the database
                //

                //
                //END of DELETING Expenses_Category record from the database.
                //

                //
                //START of clearing fields and comboBox to update comboBox with new records
                //
                /*
                comboBoxEmployees.Items.Clear();
                comboBoxEmployees.ResetText();
                */
                //
                //END of clearing fields and comboBox to update comboBox with new records
                //
            }
            else
            {
                MessageBox.Show("Please select an entry");
            }
        }

        //
        //END Delete expenses Company name
        //


        //
        //Submit buttom expenses
        //

        private void button24_Click(object sender, EventArgs e)
        {
            richTextBox3.Clear();

            if (comboBoxExpensesCategory.Text == "")
            {
                MessageBox.Show("Please Select a Category");
            }
            else
            {
                if (comboBoxbuttonExpensesBill_Type.Text == "")
                {
                    MessageBox.Show("Please Select a Type");
                }
                else
                {
                    if (comboBoxExpensesCompanies.Text == "")
                    {
                        MessageBox.Show("Please Select a Company");
                    }
                    else
                    {
                        if (exp1.Text == "")
                        {
                            MessageBox.Show("Please insert an Info");
                        }
                        else
                        {
                            if (exp2.Text == "")
                            {
                                MessageBox.Show("Please insert The Bill Reference");
                            }
                            else
                            {
                                if (exp3.Text == "")
                                {
                                    MessageBox.Show("Please insert NET - Bill Amount");
                                }
                                else
                                {
                                    if (exp4.Text == "")
                                    {
                                        MessageBox.Show("Please insert Tax Amount");
                                    }
                                    else
                                    {
                                        if (exp5.Text == "")
                                        {
                                            MessageBox.Show("Please insert Gross Amount");
                                        }
                                        else
                                        {
                                            if (exp6.Text == "")
                                            {
                                                MessageBox.Show("Please insert Date of Billing");
                                            }
                                            else
                                            {
                                                /*
                                                   comboBoxExpensesCategory.Text
                                                   comboBoxbuttonExpensesBill_Type.Text
                                                   comboBoxExpensesCompanies.Text
                                                   exp1.Text
                                                   exp2.Text
                                                   exp3.Text
                                                   exp4.Text
                                                   exp5.Text
                                                   exp6.Text
                                                */
                                                richTextBox3.AppendText(comboBoxExpensesCategory.Text + Environment.NewLine);
                                                richTextBox3.AppendText(comboBoxbuttonExpensesBill_Type.Text + Environment.NewLine);
                                                richTextBox3.AppendText(comboBoxExpensesCompanies.Text + Environment.NewLine);
                                                richTextBox3.AppendText(exp1.Text + Environment.NewLine);
                                                richTextBox3.AppendText(exp2.Text + Environment.NewLine);
                                                richTextBox3.AppendText(exp3.Text + Environment.NewLine);
                                                richTextBox3.AppendText(exp4.Text + Environment.NewLine);
                                                richTextBox3.AppendText(exp5.Text + Environment.NewLine);
                                                richTextBox3.AppendText(exp6.Text + Environment.NewLine);













































                                                using (OleDbConnection connection = new OleDbConnection(string.Format("Provider =Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))//using connection
                                                {
                                                    using (OleDbCommand insertCommand = new OleDbCommand("INSERT INTO Expenses ([Company_EXPENSES_Name],[Category_EXPENSES_Name],[Bill_type_EXPENSES_Name],[Info],[Bill_Ref],[Bill_amount_net],[Bill_amount_after_Tax],[Bill_gross],[Date_of_Billing]) VALUES (?,?,?,?,?,?,?,?,?)", connection))//insert command
                                                    using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Expenses", connection))
                                                    {
                                                        connection.Open();


                                                        DataTable table = new DataTable();
                                                        OleDbDataAdapter adapter = new OleDbDataAdapter();
                                                        adapter.SelectCommand = selectCommand;
                                                        adapter.Fill(table);


                                                        //
                                                        //start of checking if entry exists in database
                                                        //
                                                        foreach (DataRow row in table.Rows)
                                                        {
                                                            object company = row["Company_EXPENSES_Name"];
                                                            object category = row["Category_EXPENSES_Name"];
                                                            object billType = row["Bill_type_EXPENSES_Name"];
                                                            object info = row["Info"];
                                                            object billRef = row["Bill_Ref"];
                                                            object amountNet = row["Bill_amount_net"];
                                                            object amountafterTax = row["Bill_amount_after_Tax"];
                                                            object grossBill = row["Bill_gross"];
                                                            object billDate = row["Date_of_Billing"];

                                                            string strCompany = company.ToString();
                                                            string strCategory = category.ToString();
                                                            string strBillType = billType.ToString();
                                                            string strInfo = info.ToString();
                                                            string strBillRef = billRef.ToString();
                                                            string strAmountNet = amountNet.ToString();
                                                            string strAmountafterTax = amountafterTax.ToString();
                                                            string strGrossBill = grossBill.ToString();
                                                            string strBillDate = billDate.ToString();

                                                        }

                                                        insertCommand.Parameters.AddWithValue("@Company_EXPENSES_Name", WageEmployeeID.Text);
                                                        insertCommand.Parameters.AddWithValue("@Category_EXPENSES_Name", ComboBoxWage.Text);
                                                        insertCommand.Parameters.AddWithValue("@Bill_type_EXPENSES_Name", WageComboBox1.Text);
                                                        insertCommand.Parameters.AddWithValue("@Info", WageComboBox2.Text);
                                                        insertCommand.Parameters.AddWithValue("@Bill_Ref", Wage1.Text);
                                                        insertCommand.Parameters.AddWithValue("@Bill_amount_net", .Text);
                                                        insertCommand.Parameters.AddWithValue("@Bill_amount_after_Tax", .Text);
                                                        insertCommand.Parameters.AddWithValue("@Bill_gross", .Text);
                                                        insertCommand.Parameters.AddWithValue("@Date_of_Billing", .Text);


                                                        insertCommand.ExecuteNonQuery();

                                                        string strComboNames = Emp1.Text + " " + Emp2.Text + " " + Emp3.Text;

                                                        //
                                                        //end of inserting into database
                                                        //

                                                        //
                                                        //START of Showing what has been added to the database
                                                        //
                                                        richTextBox3.AppendText("Details of the New Record Added :    " + Environment.NewLine);
                                                        richTextBox3.AppendText(Environment.NewLine);
                                                        richTextBox3.AppendText(WageEmployeeID.Text + "    " + Environment.NewLine);
                                                        richTextBox3.AppendText(ComboBoxWage.Text + "    " + Environment.NewLine);
                                                        richTextBox3.AppendText(WageComboBox1.Text + "    " + Environment.NewLine);
                                                        richTextBox3.AppendText(WageComboBox2.Text + "    " + Environment.NewLine);
                                                        richTextBox3.AppendText(Wage1.Text + "    " + Environment.NewLine);
                                                        richTextBox3.AppendText(Wage2.Text + "    " + Environment.NewLine);
                                                        //
                                                        //END of Showing what has been added to the database
                                                        //

                                                        //
                                                        //end of  inserting into the database.
                                                        //

                                                    }
                                                }





























































                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        //
        //END of Submit buttom expenses
        //


        //***************************************************************************************************************************************************************************************************************************
        //
        //
        //
        //
        //
        //END of EXPENSES
        //
        //
        //
        //
        //
        //***************************************************************************************************************************************************************************************************************************


        //***************************************************************************************************************************************************************************************************************************
        //
        //
        //
        //Start of Wages 
        //
        //
        //
        //***************************************************************************************************************************************************************************************************************************

        //
        //Start of Wage comboBox.
        //

        private void ComboBoxWage_SelectedIndexChanged(object sender, EventArgs e)
        {

            //
            //Selecting Employee in the comboBox and collecting its ID and displying in the WageID txt field.
            //
            string selected = this.ComboBoxWage.GetItemText(this.ComboBoxWage.SelectedItem);
            string strWageComboEmployee = selected;

            //richTextBoxWage.AppendText(strWageComboEmployee + Environment.NewLine);

            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
            {
                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Employees", connection))
                {
                    connection.Open();

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);


                    string strComboNames = ComboBoxWage.Text; // string (Employee) inside combo box... employee selected


                    //richTextBoxEmp.AppendText(strComboBox);

                    foreach (DataRow row in table.Rows)
                    {
                        //
                        //Collecting ID from employee selec
                        //
                        object ID = row["ID"];
                        object ComboNames = row["Combo_Names"];

                        if (ComboNames.ToString() == strComboNames)
                        {
                            WageEmployeeID.Clear();
                            WageEmployeeID.AppendText(ID.ToString());
                        }
                    }
                }
            }
        }

        //
        //End of Wage comboBox.
        //

        //
        //Start of adding wages entries to the database.
        //

        private void button2_Click_2(object sender, EventArgs e)
        {
            richTextBoxWage.Clear();

            List<string> ListOfEntry = new List<string>();
            string strListOfEntryDateFrom = String.Empty;
            string strListOfEntryDateTo = String.Empty;
            //Creating a string out of the entries in the all fields to check against the database for duplicates
            string strDateFromComboBox = WageComboBox1.Text + " 00:00:00";
            string strDateToComboBox = WageComboBox2.Text + " 00:00:00";


            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider =Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))//using connection
            {
                using (OleDbCommand insertCommand = new OleDbCommand("INSERT INTO Employees_Wages ([EmployeeID],[Employee_Full_Name],[Date_From],[Date_To],[Total_Hours],[Total_Before_Tax]) VALUES (?,?,?,?,?,?)", connection))//insert command
                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Employees_Wages", connection))
                {
                    connection.Open();


                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    int iEntryCount = 0;

                    //
                    //start of checking if entry exists in database
                    //

                    foreach (DataRow row in table.Rows)
                    {
                        object ID = row["EmployeeID"];
                        object FullName = row["Employee_Full_Name"];
                        object DateFrom = row["Date_From"];
                        object DateTo = row["Date_To"];
                        object TotalHours = row["Total_Hours"];
                        object TotalBeforeTax = row["Total_Before_Tax"];

                        string strID = ID.ToString();
                        string strFullName = FullName.ToString();
                        string strDateFrom = DateFrom.ToString();
                        string strDateTo = DateTo.ToString();
                        string strTotalHours = TotalHours.ToString();
                        string strTotalBeforeTax = TotalBeforeTax.ToString();

                        ////Start of Adding all entry to a list to check if the entry exists
                        strListOfEntryDateFrom = DateFrom.ToString();
                        strListOfEntryDateTo = DateTo.ToString();

                        if (!ListOfEntry.Contains(strListOfEntryDateFrom))
                        {
                            ListOfEntry.Add(strListOfEntryDateFrom);
                        }
                        if (!ListOfEntry.Contains(strListOfEntryDateTo))
                        {
                            ListOfEntry.Add(strListOfEntryDateTo);
                        }
                        //End of Adding all entry to a list

                    }

                    //Start of Adding all items from the database into a list in order to check if new entry already exists.
                    foreach (string ListOfEntries in ListOfEntry)
                    {
                        richTextBoxWage.AppendText("-----------------------------------------------------------------------------------------" + Environment.NewLine);
                        richTextBoxWage.AppendText(ListOfEntries + "    " + Environment.NewLine);
                    }
                    // End of Adding all items from the database into a list in order to check if new entry already exists.

                    //checking if entry already exists
                    if (ListOfEntry.Contains(strDateFromComboBox) | ListOfEntry.Contains(strDateToComboBox))
                    {
                        MessageBox.Show("Date already exist Please choose another date or adit the existing entry");
                    }
                    else //if does not exists  insert into the database.
                    {

                        iEntryCount++;// To stop the loop and save only 1 time the new entry

                        if (iEntryCount == 1)
                        {
                            insertCommand.Parameters.AddWithValue("@EmployeeID", WageEmployeeID.Text);
                            insertCommand.Parameters.AddWithValue("@Employee_Full_Name", ComboBoxWage.Text);
                            insertCommand.Parameters.AddWithValue("@Date_From", WageComboBox1.Text);
                            insertCommand.Parameters.AddWithValue("@Date_To", WageComboBox2.Text);
                            insertCommand.Parameters.AddWithValue("@Total_Hours", Wage1.Text);
                            insertCommand.Parameters.AddWithValue("@Total_Before_Tax", Wage2.Text);

                            insertCommand.ExecuteNonQuery();

                            string strComboNames = Emp1.Text + " " + Emp2.Text + " " + Emp3.Text;

                            //
                            //end of inserting into database
                            //

                            //
                            //START of Showing what has been added to the database
                            //
                            richTextBoxWage.AppendText("Details of the New Record Added :    " + Environment.NewLine);
                            richTextBoxWage.AppendText(Environment.NewLine);
                            richTextBoxWage.AppendText(WageEmployeeID.Text + "    " + Environment.NewLine);
                            richTextBoxWage.AppendText(ComboBoxWage.Text + "    " + Environment.NewLine);
                            richTextBoxWage.AppendText(WageComboBox1.Text + "    " + Environment.NewLine);
                            richTextBoxWage.AppendText(WageComboBox2.Text + "    " + Environment.NewLine);
                            richTextBoxWage.AppendText(Wage1.Text + "    " + Environment.NewLine);
                            richTextBoxWage.AppendText(Wage2.Text + "    " + Environment.NewLine);
                            //
                            //END of Showing what has been added to the database
                            //

                            //
                            //end of  inserting into the database.
                            //

                        }
                    }
                }
            }
        }

        //
        //End of adding wages entries to the database.
        //

        //***************************************************************************************************************************************************************************************************************************
        //
        //
        //
        //End of Wages 
        //
        //
        //
        //***************************************************************************************************************************************************************************************************************************


        private void button1_Click(object sender, EventArgs e)
        { }
        private void button2_Click(object sender, EventArgs e)
        { }
        private void tabPage1_Click(object sender, EventArgs e)
        { }
        private void tabPage2_Click(object sender, EventArgs e)
        { }
        private void tabPage3_Click(object sender, EventArgs e)
        { }
        private void panel1_Paint(object sender, PaintEventArgs e)
        { }
        private void button5_Click(object sender, EventArgs e)
        { }
        private void button2_Click_1(object sender, EventArgs e)
        { }
        private void panel2_Paint(object sender, PaintEventArgs e)
        { }
        private void label15_Click(object sender, EventArgs e)
        { }
        private void textBox12_TextChanged(object sender, EventArgs e)
        { }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        { }
        private void label15_Click_1(object sender, EventArgs e)
        { }
        private void label18_Click(object sender, EventArgs e)
        { }
        private void tabPage4_Click(object sender, EventArgs e)
        { }
        private void label17_Click(object sender, EventArgs e)
        { }
        private void label23_Click(object sender, EventArgs e)
        { }
        private void label20_Click(object sender, EventArgs e)
        { }
        private void label25_Click(object sender, EventArgs e)
        { }
        private void label22_Click(object sender, EventArgs e)
        { }
        private void label18_Click_1(object sender, EventArgs e)
        { }
        private void label21_Click(object sender, EventArgs e)
        { }
        private void label24_Click(object sender, EventArgs e)
        { }
        private void label26_Click(object sender, EventArgs e)
        { }
        private void label24_Click_1(object sender, EventArgs e)
        { }
        private void label23_Click_1(object sender, EventArgs e)
        { }
        private void label28_Click(object sender, EventArgs e)
        { }
        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        { }
        private void panel4_Paint(object sender, PaintEventArgs e)
        { }
        private void button19_Click(object sender, EventArgs e)
        { }
        private void label14_Click(object sender, EventArgs e)
        { }
        private void panel8_Paint(object sender, PaintEventArgs e)
        { }
        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        { }
        private void label16_Click(object sender, EventArgs e)
        { }
        private void EmpID_TextChanged(object sender, EventArgs e)
        { }
        private void comboBoxbuttonExpensesBill_Type_SelectedIndexChanged(object sender, EventArgs e)
        { }
        private void comboBoxExpensesCompanies_SelectedIndexChanged(object sender, EventArgs e)
        { }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }
    }
}
