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

            //
            //select command
            //
            //string mdfFile1 = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\IsaqueH\Documents\GDB.mdb";


            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
            {
                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Employees", connection))
                {
                    connection.Open();

                    List<string> ListOfNames = new List<string>();

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

            //***********************************************************************************************************************************************************************************************************************
            //
            //
            //START of Expenses category combo box On Load -  ComboBox to show Expenses category, Bill type & Company name from the database
            //
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
                    List<string> ListOfCategory_EXPENSES_Name = new List<string>();

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        object CategoryNameValue = row["Category_EXPENSES_Name"];
                        string strExpensesCategoryName = CategoryNameValue + "";//

                        if (!ListOfCategory_EXPENSES_Name.Contains(strExpensesCategoryName))
                        {
                            comboBoxExpensesCategory.Items.Add(strExpensesCategoryName);//adding items to the comboBox Expenses category Tab
                            ListOfCategory_EXPENSES_Name.Add(strExpensesCategoryName);
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
                    List<string> ListOfBill_type_EXPENSES_Name = new List<string>();

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        object BillTypeNameValue = row["Bill_type_EXPENSES_Name"];
                        string strExpensesBillTypeName = BillTypeNameValue + "";//

                        if (!ListOfBill_type_EXPENSES_Name.Contains(strExpensesBillTypeName))
                        {
                            comboBoxbuttonExpensesBill_Type.Items.Add(strExpensesBillTypeName);//adding items to the comboBox Expenses category Tab
                            ListOfBill_type_EXPENSES_Name.Add(strExpensesBillTypeName);
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
                    List<string> ListOfCompany_EXPENSES_Name = new List<string>();

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        object CategoryNameValue = row["Company_EXPENSES_Name"];
                        string strExpensesCompanyName = CategoryNameValue + "";//

                        if (!ListOfCompany_EXPENSES_Name.Contains(strExpensesCompanyName))
                        {
                            comboBoxExpensesCompanies.Items.Add(strExpensesCompanyName);//adding items to the comboBox Expenses category Tab
                            ListOfCompany_EXPENSES_Name.Add(strExpensesCompanyName);
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



            //***************************************************************************************************************************************************************************************************************************
            //
            //
            //
            //END of On Load -  showing Expenses category, Bill type & Company name when the program starts
            //
            //
            //
            //***************************************************************************************************************************************************************************************************************************

            //########


            //***********************************************************************************************************************************************************************************************************************
            //
            //
            //START of On Load Income category combo box On Load -  ComboBox to show Expenses Income, Bill type & Company name from the database
            //
            //
            //***********************************************************************************************************************************************************************************************************************

            //START of Updating ComboBox to show Income category
            //
            //
            //select command
            //



            comboBoxIncomeCategory.AutoCompleteMode = AutoCompleteMode.Suggest;//ComboBox Auto suggestion when typing first characters 
            comboBoxIncomeCategory.AutoCompleteSource = AutoCompleteSource.ListItems;//ComboBox Auto suggestion when typing first characters 



            comboBoxIncomeCategory.Sorted = true;

            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
            {
                List<string> ListOfIncome_Category = new List<string>();
                //
                //collecting from Income
                //
                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Income_Category", connection))
                {
                    connection.Open();

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        object CategoryNameValue = row["Category"];
                        string strIncomeCategoryName = CategoryNameValue + "";//

                        if (!ListOfIncome_Category.Contains(strIncomeCategoryName))
                        {
                            comboBoxIncomeCategory.Items.Add(strIncomeCategoryName);//adding items to the comboBox Income category Tab
                            ListOfIncome_Category.Add(strIncomeCategoryName);
                        }

                    }
                }
                //
                //End of collecting from Income Category
                //
            }
            //***********************************************************************************************************************************************************************************************************************
            //
            //
            //END of Income On Loadcategory combo box On Load -  ComboBox to show Expenses Income, Bill type & Company name from the database
            //
            //
            //
            //***********************************************************************************************************************************************************************************************************************

            //########


            //***********************************************************************************************************************************************************************************************************************
            //
            //
            //START of On Load Income Type_Of_Sale combo box On Load -  ComboBox to show Income, Bill type & Company name from the database
            //
            //
            //***********************************************************************************************************************************************************************************************************************

            //START of Updating ComboBox to show Income Type_Of_Sale
            //
            //
            //select command
            //



            comboBoxIncomeType_Of_Sale.AutoCompleteMode = AutoCompleteMode.Suggest;//ComboBox Auto suggestion when typing first characters 
            comboBoxIncomeType_Of_Sale.AutoCompleteSource = AutoCompleteSource.ListItems;//ComboBox Auto suggestion when typing first characters 



            comboBoxIncomeType_Of_Sale.Sorted = true;

            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
            {
                List<string> ListOfIncome_Type_Of_Sale = new List<string>();
                //
                //collecting from Income
                //
                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Income_Type_Of_Sale", connection))
                {
                    connection.Open();

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        object Type_Of_SaleNameValue = row["Type_Of_Sale"];
                        string strIncomeType_Of_Sale = Type_Of_SaleNameValue + "";//

                        if (!ListOfIncome_Type_Of_Sale.Contains(strIncomeType_Of_Sale))
                        {
                            comboBoxIncomeType_Of_Sale.Items.Add(strIncomeType_Of_Sale);//adding items to the comboBox Income Type_Of_Sale Tab
                            ListOfIncome_Type_Of_Sale.Add(strIncomeType_Of_Sale);
                        }

                    }
                }
                //
                //End of collecting from Income Type_Of_Sale
                //
            }
            //***********************************************************************************************************************************************************************************************************************
            //
            //
            //END of Income On Load Type_Of_Sale combo box On Load -  ComboBox to show Income, Bill type & Company name from the database
            //
            //
            //
            //***********************************************************************************************************************************************************************************************************************


            //***********************************************************************************************************************************************************************************************************************
            //
            //
            //START of On Load Income Payment_Method combo box On Load -  ComboBox to show Expenses Income, Bill type & Company name from the database
            //
            //
            //***********************************************************************************************************************************************************************************************************************

            //START of Updating ComboBox to show Income Payment_Method
            //
            //
            //select command
            //



            comboBoxIncomePayment_Method.AutoCompleteMode = AutoCompleteMode.Suggest;//ComboBox Auto suggestion when typing first characters 
            comboBoxIncomePayment_Method.AutoCompleteSource = AutoCompleteSource.ListItems;//ComboBox Auto suggestion when typing first characters 



            comboBoxIncomePayment_Method.Sorted = true;

            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
            {
                List<string> ListOfIncome_Payment_Method = new List<string>();
                //
                //collecting from Income
                //
                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Income_Payment_Method", connection))
                {
                    connection.Open();

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        object Payment_MethodNameValue = row["Payment_Method"];
                        string strIncomePayment_MethodName = Payment_MethodNameValue + "";//

                        if (!ListOfIncome_Payment_Method.Contains(strIncomePayment_MethodName))
                        {
                            comboBoxIncomePayment_Method.Items.Add(strIncomePayment_MethodName);//adding items to the comboBox Income Payment_Method Tab
                            ListOfIncome_Payment_Method.Add(strIncomePayment_MethodName);
                        }

                    }
                }
                //
                //End of collecting from Income Payment_Method
                //
            }
            //***********************************************************************************************************************************************************************************************************************
            //
            //
            //END of Income On LoadPayment_Method combo box On Load -  ComboBox to show Expenses Income, Bill type & Company name from the database
            //
            //
            //
            //***********************************************************************************************************************************************************************************************************************

            //########
        }

        //***************************************************************************************************************************************************************************************************************************
        //
        //
        //
        //START On Load -  showing Expenses category, Bill type & Company name when the program starts
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
                                insertCommand.Parameters.AddWithValue("@Start_Date", Emp11.Value);
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
                                richTextBoxEmp.AppendText(Emp11.Value + "    " + Environment.NewLine);
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
            //Emp11.Clear();
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
            //Emp11.Clear();
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
                                Emp11.Value = new DateTime(Convert.ToInt64(Convert.ToDecimal(StartDate)));// converting lo long

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
                    updateCommand.Parameters.AddWithValue("@Start_Date", Emp11.Value);

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
                    richTextBoxEmp.AppendText(Emp11.Value + "    " + Environment.NewLine);
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
            //Emp11.Clear();
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
            richTextBoxEmp.AppendText(Emp11.Value + "    " + Environment.NewLine);
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
            //Emp11.Clear();
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


        //######



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
        //START of Delete expenses category
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
                                            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider =Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))//using connection
                                            {
                                                using (OleDbCommand insertCommand = new OleDbCommand("INSERT INTO Expenses ([Company_EXPENSES_Name],[Category_EXPENSES_Name],[Bill_type_EXPENSES_Name],[Info],[Bill_Ref],[Bill_amount_net],[Bill_amount_after_Tax],[Bill_gross],[Date_of_Billing]) VALUES (?,?,?,?,?,?,?,?,?)", connection))//insert command
                                                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Expenses", connection))
                                                {
                                                    connection.Open();

                                                    string strBillRef = string.Empty;
                                                    List<string> ListBillRef = new List<string>();

                                                    DataTable table = new DataTable();
                                                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                                                    adapter.SelectCommand = selectCommand;
                                                    adapter.Fill(table);


                                                    //
                                                    //start of checking if entry exists in database
                                                    //

                                                    foreach (DataRow row in table.Rows)
                                                    {
                                                        object billRef = row["Bill_Ref"];

                                                        strBillRef = billRef.ToString();
                                                        //adding to a list
                                                        ListBillRef.Add(strBillRef);
                                                    }

                                                    if (ListBillRef.Contains(Emp2.Text))
                                                    {
                                                        MessageBox.Show("Expense with this reference already exists");

                                                    }
                                                    else
                                                    {
                                                        insertCommand.Parameters.AddWithValue("@Company_EXPENSES_Name", comboBoxExpensesCategory.Text);
                                                        insertCommand.Parameters.AddWithValue("@Category_EXPENSES_Name", comboBoxbuttonExpensesBill_Type.Text);
                                                        insertCommand.Parameters.AddWithValue("@Bill_type_EXPENSES_Name", comboBoxExpensesCompanies.Text);
                                                        insertCommand.Parameters.AddWithValue("@Info", exp1.Text);
                                                        insertCommand.Parameters.AddWithValue("@Bill_Ref", exp2.Text);
                                                        insertCommand.Parameters.AddWithValue("@Bill_amount_net", exp3.Text);
                                                        insertCommand.Parameters.AddWithValue("@Bill_amount_after_Tax", exp4.Text);
                                                        insertCommand.Parameters.AddWithValue("@Bill_gross", exp5.Text);
                                                        insertCommand.Parameters.AddWithValue("@Date_of_Billing", exp6.Value);


                                                        insertCommand.ExecuteNonQuery();


                                                        //
                                                        //end of inserting into database
                                                        //

                                                        //
                                                        //START of Showing what has been added to the database
                                                        //
                                                        richTextBox3.AppendText("Details of the New Expense Record Added :    " + Environment.NewLine);
                                                        richTextBox3.AppendText(Environment.NewLine);
                                                        richTextBox3.AppendText(comboBoxExpensesCategory.Text + Environment.NewLine);
                                                        richTextBox3.AppendText(comboBoxbuttonExpensesBill_Type.Text + Environment.NewLine);
                                                        richTextBox3.AppendText(comboBoxExpensesCompanies.Text + Environment.NewLine);
                                                        richTextBox3.AppendText(exp1.Text + Environment.NewLine);
                                                        richTextBox3.AppendText(exp2.Text + Environment.NewLine);
                                                        richTextBox3.AppendText(exp3.Text + Environment.NewLine);
                                                        richTextBox3.AppendText(exp4.Text + Environment.NewLine);
                                                        richTextBox3.AppendText(exp5.Text + Environment.NewLine);
                                                        richTextBox3.AppendText(exp6.Value + Environment.NewLine);
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

        //
        //Start of Selecting expenses records
        //
        
        private void button26_Click(object sender, EventArgs e)
        { }


        //
        //
        //Start of Selecting expenses records
        //
        //

        private void button1_Click_1(object sender, EventArgs e)
        {
            using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
            {
                using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Expenses", connection))
                {
                    connection.Open();

                    List<string> ListOfNames = new List<string>();
                    int icount = 0;
                    int icount2 = 0;

                    DataTable table = new DataTable();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = selectCommand;
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                    {
                        // collect data here.
                        icount++;
                    }
                    
                    
                    foreach (DataRow row in table.Rows)
                    {
                        icount2++;
                        if (icount2 == icount - 4 || icount2 == icount - 3 || icount2 == icount - 2 || icount2 == icount -1 || icount2 == icount)
                        {
                            
                            object CompanyExpensesName = row["Company_EXPENSES_Name"];
                            object CategoryExpensesName = row["Category_EXPENSES_Name"];
                            object BillTypeExpensesName = row["Bill_type_EXPENSES_Name"];
                            object ExpensesInfo = row["Info"];
                            object ExpensesBillRef = row["Bill_Ref"];
                            object ExpensesBillAmoutNet = row["Bill_amount_net"];
                            object ExpensesBillAfetTax = row["Bill_amount_after_Tax"];
                            object ExpensesBillGross = row["Bill_gross"];
                            object ExpensesDateofBilling = row["Date_of_Billing"];

                            string strIncomeCategoryName = CompanyExpensesName + "";
                            string strCategoryExpensesName = CategoryExpensesName + "";
                            string strBillTypeExpensesName = BillTypeExpensesName + "";
                            string strExpensesInfo = ExpensesInfo + "";
                            string strExpensesBillRef = ExpensesBillRef + "";
                            string strExpensesBillAmoutNet = ExpensesBillAmoutNet + "";
                            string strExpensesBillAfetTax = ExpensesBillAfetTax + "";
                            string strExpensesBillGross = ExpensesBillGross + "";
                            string strExpensesDateofBilling = ExpensesDateofBilling + "";

                            richTextBox3.AppendText("Details of the Last 5 Expenses Record Added :    " + Environment.NewLine);
                            richTextBox3.AppendText("Record " + icount2 + "  -  " + strIncomeCategoryName + " " + strCategoryExpensesName + " " + BillTypeExpensesName  + " " + strExpensesInfo + " " + strExpensesBillRef + " " + strExpensesBillAmoutNet + " " + strExpensesBillAfetTax + " " + strExpensesBillGross + "  " + strExpensesDateofBilling + Environment.NewLine);
                            richTextBox3.AppendText(Environment.NewLine);

                        }
                    }
                }
            }
        }

        //
        //
        //End of Selecting expenses records
        //
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

        //########

        //***************************************************************************************************************************************************************************************************************************
        //
        //
        //
        //Start of Incomes
        //
        //
        //
        //***************************************************************************************************************************************************************************************************************************

        //
        //Category ComboBox
        //
        private void comboBoxIncomeCategory_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void incomebutton1_Click(object sender, EventArgs e)
        {

            //
            //START of Delete Income category
            //

            incomerichTextBox.Clear();
            //
            //START of DELETING Income_Category  record from the database.
            //            

            if (comboBoxIncomeCategory.Text != "")
            {
                using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
                {
                    using (OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM Income_Category WHERE [Category] = ?", connection))
                    {
                        connection.Open();

                        deleteCommand.Parameters.AddWithValue("@Category", comboBoxIncomeCategory.Text);

                        deleteCommand.ExecuteNonQuery();
                    }
                }

                //
                //START of Showing what has been updated into the database
                //
                incomerichTextBox.AppendText("Records Deleted Successfully:    " + Environment.NewLine);
                incomerichTextBox.AppendText(comboBoxIncomeCategory.Text + Environment.NewLine);
                comboBoxIncomeCategory.Items.Remove(comboBoxIncomeCategory.Text);


                //
                //END of Showing what has been updated into the database
                //

                //
                //END of DELETING Income_Category record from the database.
                //

                //
                //START of clearing fields and comboBox to update comboBox with new records
                //
                
                //comboBoxIncomeCategory.Items.Clear();
                //comboBoxIncomeCategory.ResetText();
                
                //
                //END of clearing fields and comboBox to update comboBox with new records
                //
            }
            else
            {
                MessageBox.Show("Please select an entry");
            }
        }

        private void incomebutton2_Click(object sender, EventArgs e)
        {
            //
            //Adding a new category Income to the database using a input message box
            //
            List<string> List_Category = new List<string>();// to add all the categories to a list for latter to be checked if new entry already exists in the table
            string NewCategoryIncomeValue = Interaction.InputBox("Please insert a new category", "New Category", "");

            incomerichTextBox.AppendText(NewCategoryIncomeValue);

            //
            //End of collecting a new category Income to add in the database using a inputBox message box
            //
            if (NewCategoryIncomeValue != "" && comboBoxIncomeCategory.Text == "")
            {

                incomerichTextBox.Clear();

                //
                //START of new records to the Income_Category Table.
                //
                using (OleDbConnection connection = new OleDbConnection(string.Format("Provider =Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))//using connection
                {

                    using (OleDbCommand insertCommand = new OleDbCommand("INSERT INTO Income_Category ([Category]) VALUES (?)", connection))//insert command
                    using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Income_Category", connection))
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

                            object nameValue = row["Category"];

                            string strName = nameValue.ToString();

                            if (!List_Category.Contains(strName))
                            {
                                List_Category.Add(strName);
                            }

                            //
                            //End of checking if entry exists in database
                            //
                        }

                        if (List_Category.Contains(NewCategoryIncomeValue))
                        {
                            MessageBox.Show("Category already exist");
                        }
                        //
                        //inserting into the database
                        //
                        else
                        {
                            insertCommand.Parameters.AddWithValue("Category", NewCategoryIncomeValue);

                            insertCommand.ExecuteNonQuery();

                            connection.Close();

                            //
                            //end of inserting into database
                            //

                            //
                            //START of Showing what has been added to the database
                            //
                            incomerichTextBox.AppendText("Records Inserted:    " + Environment.NewLine);
                            incomerichTextBox.AppendText(Environment.NewLine);
                            incomerichTextBox.AppendText(NewCategoryIncomeValue + "    " + Environment.NewLine);
                            //
                            //END of Showing what has been added to the database
                            //

                            //
                            //adding to the comboBox
                            //

                            comboBoxIncomeCategory.Items.Add(NewCategoryIncomeValue);

                            //
                            //end of adding to the comboBox
                            //
                        }



                    }
                }
                //
                //END of Adding a new Income category
                //

                //
                //START of clearing fields and comboBox to update comboBox with new records
                //
                /*
                comboBoxIncomeCategory.Items.Clear();
                comboBoxIncomeCategory.ResetText();
                */
                //
                //END of clearing fields and comboBox to update comboBox with new records
                //

                // END of on click to check for an existing entri and add a new entry to categories Income
            }
        }


        private void incomebutton5_Click_1(object sender, EventArgs e)
        {

            //
            //START of Delete Income Payment_Method
            //

            incomerichTextBox.Clear();
            //
            //START of DELETING Income_Payment_Method  record from the database.
            //            

            if (comboBoxIncomePayment_Method.Text != "")
            {
                using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
                {
                    using (OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM Income_Payment_Method WHERE [Payment_Method] = ?", connection))
                    {
                        connection.Open();

                        deleteCommand.Parameters.AddWithValue("@Payment_Method", comboBoxIncomePayment_Method.Text);

                        deleteCommand.ExecuteNonQuery();
                    }
                }

                //
                //START of Showing what has been updated into the database
                //
                incomerichTextBox.AppendText("Records Deleted Successfully:    " + Environment.NewLine);
                incomerichTextBox.AppendText(comboBoxIncomePayment_Method.Text + Environment.NewLine);
                comboBoxIncomePayment_Method.Items.Remove(comboBoxIncomePayment_Method.Text);


                //
                //END of Showing what has been updated into the database
                //

                //
                //END of DELETING Income_Payment_Method record from the database.
                //

                //
                //START of clearing fields and comboBox to update comboBox with new records
                //

                //comboBoxIncomePayment_Method.Items.Clear();
                //comboBoxIncomePayment_Method.ResetText();

                //
                //END of clearing fields and comboBox to update comboBox with new records
                //
            }
            else
            {
                MessageBox.Show("Please select an entry");
            }
        }
        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void incomebutton3_Click(object sender, EventArgs e)
        {

            //
            //START of Delete Income Type_Of_Sale
            //

            incomerichTextBox.Clear();
            //
            //START of DELETING Income_Type_Of_Sale  record from the database.
            //            

            if (comboBoxIncomeType_Of_Sale.Text != "")
            {
                using (OleDbConnection connection = new OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))
                {
                    using (OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM Income_Type_Of_Sale WHERE [Type_Of_Sale] = ?", connection))
                    {
                        connection.Open();

                        deleteCommand.Parameters.AddWithValue("@Type_Of_Sale", comboBoxIncomeType_Of_Sale.Text);

                        deleteCommand.ExecuteNonQuery();
                    }
                }

                //
                //START of Showing what has been updated into the database
                //
                incomerichTextBox.AppendText("Records Deleted Successfully:    " + Environment.NewLine);
                incomerichTextBox.AppendText(comboBoxIncomeType_Of_Sale.Text + Environment.NewLine);
                comboBoxIncomeType_Of_Sale.Items.Remove(comboBoxIncomeType_Of_Sale.Text);


                //
                //END of Showing what has been updated into the database
                //

                //
                //END of DELETING Income_Type_Of_Sale record from the database.
                //

                //
                //START of clearing fields and comboBox to update comboBox with new records
                //
                
                //comboBoxIncomeType_Of_Sale.Items.Clear();
                //comboBoxIncomeType_Of_Sale.ResetText();
                
                //
                //END of clearing fields and comboBox to update comboBox with new records
                //
            }
            else
            {
                MessageBox.Show("Please select an entry");
            }
        }

        private void incomebutton4_Click(object sender, EventArgs e)
        {

            //
            //Adding a new Type_Of_Sale Income to the database using a input message box
            //
            List<string> List_Type_Of_Sale = new List<string>();// to add all the Type_Of_Sale to a list for latter to be checked if new entry already exists in the table
            string NewType_Of_SaleIncomeValue = Interaction.InputBox("Please insert a new Type_Of_Sale", "New Type_Of_Sale", "");

            incomerichTextBox.AppendText(NewType_Of_SaleIncomeValue);

            //
            //End of collecting a new Type_Of_Sale Income to add in the database using a inputBox message box
            //
            if (NewType_Of_SaleIncomeValue != "" && comboBoxIncomeType_Of_Sale.Text == "")
            {

                incomerichTextBox.Clear();

                //
                //START of new records to the Income_Type_Of_Sale Table.
                //
                using (OleDbConnection connection = new OleDbConnection(string.Format("Provider =Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))//using connection
                {

                    using (OleDbCommand insertCommand = new OleDbCommand("INSERT INTO Income_Type_Of_Sale ([Type_Of_Sale]) VALUES (?)", connection))//insert command
                    using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Income_Type_Of_Sale", connection))
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

                            object nameValue = row["Type_Of_Sale"];

                            string strName = nameValue.ToString();

                            if (!List_Type_Of_Sale.Contains(strName))
                            {
                                List_Type_Of_Sale.Add(strName);
                            }

                            //
                            //End of checking if entry exists in database
                            //
                        }

                        if (List_Type_Of_Sale.Contains(NewType_Of_SaleIncomeValue))
                        {
                            MessageBox.Show("Type_Of_Sale already exist");
                        }
                        //
                        //inserting into the database
                        //
                        else
                        {
                            insertCommand.Parameters.AddWithValue("Type_Of_Sale", NewType_Of_SaleIncomeValue);

                            insertCommand.ExecuteNonQuery();

                            connection.Close();

                            //
                            //end of inserting into database
                            //

                            //
                            //START of Showing what has been added to the database
                            //
                            incomerichTextBox.AppendText("Records Inserted:    " + Environment.NewLine);
                            incomerichTextBox.AppendText(Environment.NewLine);
                            incomerichTextBox.AppendText(NewType_Of_SaleIncomeValue + "    " + Environment.NewLine);
                            //
                            //END of Showing what has been added to the database
                            //

                            //
                            //adding to the comboBox
                            //

                            comboBoxIncomeType_Of_Sale.Items.Add(NewType_Of_SaleIncomeValue);

                            //
                            //end of adding to the comboBox
                            //
                        }



                    }
                }
                //
                //END of Adding a new Income Type_Of_Sale
                //

                //
                //START of clearing fields and comboBox to update comboBox with new records
                //
                /*
                comboBoxIncomeType_Of_Sale.Items.Clear();
                comboBoxIncomeType_Of_Sale.ResetText();
                */
                //
                //END of clearing fields and comboBox to update comboBox with new records
                //

                // END of on click to check for an existing entri and add a new entry to Type_Of_Sale Income
            }
        }

        private void incomebutton5_Click(object sender, EventArgs e)
        {

            //
            //Adding a new Payment_Method Income to the database using a input message box
            //
            List<string> List_Payment_Method = new List<string>();// to add all the Payment_Method to a list for latter to be checked if new entry already exists in the table
            string NewPayment_MethodIncomeValue = Interaction.InputBox("Please insert a new Payment_Method", "New Payment_Method", "");

            incomerichTextBox.AppendText(NewPayment_MethodIncomeValue);

            //
            //End of collecting a new Payment_Method Income to add in the database using a inputBox message box
            //
            if (NewPayment_MethodIncomeValue != "" && comboBoxIncomePayment_Method.Text == "")
            {

                incomerichTextBox.Clear();

                //
                //START of new records to the Income_Payment_Method Table.
                //
                using (OleDbConnection connection = new OleDbConnection(string.Format("Provider =Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))//using connection
                {

                    using (OleDbCommand insertCommand = new OleDbCommand("INSERT INTO Income_Payment_Method ([Payment_Method]) VALUES (?)", connection))//insert command
                    using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Income_Payment_Method", connection))
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

                            object nameValue = row["Payment_Method"];

                            string strName = nameValue.ToString();

                            if (!List_Payment_Method.Contains(strName))
                            {
                                List_Payment_Method.Add(strName);
                            }

                            //
                            //End of checking if entry exists in database
                            //
                        }

                        if (List_Payment_Method.Contains(NewPayment_MethodIncomeValue))
                        {
                            MessageBox.Show("Payment_Method already exist");
                        }
                        //
                        //inserting into the database
                        //
                        else
                        {
                            insertCommand.Parameters.AddWithValue("Payment_Method", NewPayment_MethodIncomeValue);

                            insertCommand.ExecuteNonQuery();

                            connection.Close();

                            //
                            //end of inserting into database
                            //

                            //
                            //START of Showing what has been added to the database
                            //
                            incomerichTextBox.AppendText("Records Inserted:    " + Environment.NewLine);
                            incomerichTextBox.AppendText(Environment.NewLine);
                            incomerichTextBox.AppendText(NewPayment_MethodIncomeValue + "    " + Environment.NewLine);
                            //
                            //END of Showing what has been added to the database
                            //

                            //
                            //adding to the comboBox
                            //

                            comboBoxIncomePayment_Method.Items.Add(NewPayment_MethodIncomeValue);

                            //
                            //end of adding to the comboBox
                            //
                        }



                    }
                }
                //
                //END of Adding a new Income Payment_Method
                //

                //
                //START of clearing fields and comboBox to update comboBox with new records
                //
                /*
                comboBoxIncomePayment_Method.Items.Clear();
                comboBoxIncomePayment_Method.ResetText();
                */
                //
                //END of clearing fields and comboBox to update comboBox with new records
                //

                // END of on click to check for an existing entri and add a new entry to Payment_Method Income
            }
        }
        // Submit
        private void incomebutton6_Click(object sender, EventArgs e)
        {
            
            //
            //
            //START of Adding a new Employee
            //
            //

            incomerichTextBox.Clear();
            //
            //inserting into the database
            //
            if (incTxt1.Text == "" || incDtm.Text == "" || incTxt2.Text == "")
            {
                MessageBox.Show("Please insert data to all fields");

            }
            else
            {
                using (OleDbConnection connection = new OleDbConnection(string.Format("Provider =Microsoft.Jet.OLEDB.4.0;Data Source={0}", mdfFile)))//using connection
                {

                    using (OleDbCommand insertCommand = new OleDbCommand("INSERT INTO Income ([Category],[Type_Of_Sale],[Info],[Payment_Method],[Date1],[Amount]) VALUES (?,?,?,?,?,?)", connection))//insert command
                    using (OleDbCommand selectCommand = new OleDbCommand("SELECT * FROM Income", connection))
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



                        //
                        //inserting into the database
                        //
                        /*
                        insertCommand.Parameters.AddWithValue("@Category", comboBoxIncomeCategory.Text);
                        insertCommand.Parameters.AddWithValue("@Type_Of_Sale", comboBoxIncomeType_Of_Sale.Text);
                        insertCommand.Parameters.AddWithValue("@Info", incTxt1.Text);
                        insertCommand.Parameters.AddWithValue("@Amount", incTxt2.Text);
                        insertCommand.Parameters.AddWithValue("@Payment_Method", comboBoxIncomePayment_Method.Text);
                        insertCommand.Parameters.AddWithValue("@Date1", incDtm.Value);


                        insertCommand.ExecuteNonQuery();
                        */
                        //
                        //end of inserting into database
                        //

                        //
                        //START of Showing what has been added to the database
                        //
                        incomerichTextBox.AppendText("Details of the New Record Added :    " + Environment.NewLine);
                        incomerichTextBox.AppendText(Environment.NewLine);
                        incomerichTextBox.AppendText(comboBoxIncomeCategory.Text + "    " + Environment.NewLine);
                        incomerichTextBox.AppendText(comboBoxIncomeType_Of_Sale.Text + "    " + Environment.NewLine);
                        incomerichTextBox.AppendText(incTxt1.Text + "    " + Environment.NewLine);
                        incomerichTextBox.AppendText(incDtm.Value + "    " + Environment.NewLine);
                        incomerichTextBox.AppendText(comboBoxIncomePayment_Method.Text + "    " + Environment.NewLine);
                        incomerichTextBox.AppendText(incTxt2.Text + "    " + Environment.NewLine);
                        //
                        //END of Showing what has been added to the database
                        //
                    }
                    
                }
                
            }
            //
            //END of Adding a new Income
            //
            
        }


        //***************************************************************************************************************************************************************************************************************************
        //
        //
        //
        //End of Incomes
        //
        //
        //
        //***************************************************************************************************************************************************************************************************************************

        //########

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
            string strDateFromComboBox = WageComboBox1.Value + " 00:00:00";
            string strDateToComboBox = WageComboBox1.Value + " 00:00:00";


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
                    /*
                    foreach (string ListOfEntries in ListOfEntry)
                    {
                        richTextBoxWage.AppendText("-----------------------------------------------------------------------------------------" + Environment.NewLine);
                        richTextBoxWage.AppendText(ListOfEntries + "    " + Environment.NewLine);
                    }
                    */
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
                            insertCommand.Parameters.AddWithValue("@Date_From", WageComboBox1.Value);
                            insertCommand.Parameters.AddWithValue("@Date_To", WageComboBox1.Value);
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
                            richTextBoxWage.AppendText(WageComboBox1.Value + "    " + Environment.NewLine);
                            richTextBoxWage.AppendText(WageComboBox1.Value + "    " + Environment.NewLine);
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

        //########



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
        { }

        private void Wage5_ValueChanged(object sender, EventArgs e)
        { }

        private void exp6_ValueChanged(object sender, EventArgs e)
        { }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        { }

        
    }
}
