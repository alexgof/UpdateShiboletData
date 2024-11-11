using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

namespace UpdateShiboletData
{
    public partial class FormUpdateData : Form
    {
        //private string connectionString = "Server=ALEX-PC;Database=ShiboletDB;Integrated Security=SSPI;";// User Id=your_user;Password=your_password;";
        private string connectionString = ConfigurationManager.ConnectionStrings["MyDatabaseConnection"].ConnectionString;
        private string password = ConfigurationManager.AppSettings["FilePassword"];
        private string originalFilePath = string.Empty;
        private string destinationFilePath = ConfigurationManager.AppSettings["DestinationFilePath"];
        private string rashutId = ConfigurationManager.AppSettings["RashutId"];

        public FormUpdateData()
        {
            InitializeComponent();
        }

        private void buttonSelectFile_Click(object sender, EventArgs e)
        {
            // Create a new instance of OpenFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Set the filter for file types (optional)
            //openFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
            openFileDialog.Filter = "Text Files (*.xlsx)|*.xlsx";

            // Show the dialog and check if the user selected a file
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Get the selected file's path
                string selectedFilePath = openFileDialog.FileName;

                // Display the file path in a TextBox (or handle it as needed)
                textBoxFilePath.Text = selectedFilePath;
            }
        }

        // Button to save data from CSV to SQL Server
        private void btnSaveData_Click(object sender, EventArgs e)
        {
           if( LoadDataFromExcel())
            ReadExcelWithPasswordAndSaveToDB();
        }

        private bool LoadDataFromExcel()
        {
            originalFilePath = textBoxFilePath.Text;

            if (string.IsNullOrEmpty(originalFilePath) || !File.Exists(originalFilePath))
            {
                MessageBox.Show("Please select a valid excel file.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;        
        }

        private void ReadExcelWithoutPassword()
        {

            // Read the CSV file and insert data
            try
            {
                using (StreamReader reader = new StreamReader(originalFilePath))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        string[] values = line.Split(',');

                        // Insert the values into the database (adjust columns as needed)
                        InsertDataToDatabase(values, connectionString);

                    }
                }
                MessageBox.Show("Data has been successfully inserted into the database.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void ReadExcelWithPasswordAndSaveToDB()
        {
            FileInfo fileInfo = new FileInfo(originalFilePath);

            // Ensure ExcelPackage.LicenseContext is set to avoid license prompt
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // or LicenseContext.Commercial if you have a license

            MessageBox.Show($"Password: {password}, from file:{ originalFilePath}",Application.ProductName,MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
            using (ExcelPackage package = new ExcelPackage(fileInfo, password))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Access the first worksheet
                int counter = 0;
                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    string[] values = new string[worksheet.Dimension.End.Column];
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Text;
                        Console.WriteLine($"Row {row}, Col {col}: {cellValue}");
                        values[col - 1] = cellValue;
                    }

                    if (row>16 && !string.IsNullOrEmpty(values[4]))
                    {
                        if (Int32.TryParse(values[0], out int id))
                        {
                            //DeleteAllDataFromCurrentDay ONLY ONES
                            if (row == 17) DeleteAllDataFromCurrentDay();



                            // Insert the values into the database (adjust columns as needed)
                            InsertDataToDatabase(values, connectionString);
                            counter++;
                        }
                    }
                }
                MessageBox.Show($"Data inserted successfully. Rows count: {counter} ", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }            
        }

        // Method to insert data into SQL Server
        private void InsertDataToDatabase(string[] values, string connectionString)
        {
            //string query = "INSERT INTO shibolet_tbl (insert_date, Column1, Column2, Column3,Column4, Column5, Column6,Column7, Column8, Column9,Column10) " +
            //    "VALUES (@insert_date, @Value1, @Value2, @Value3,@Value4, @Value5, @Value6,@Value7, @Value8, @Value9,@Value10)";
            string query = "INSERT INTO shibolet_tbl (insert_date, file_index,id_number,id_type,full_name,age,age_group,gender,nationality,number_of_years_in_the_country,country_of_immigration,tlm,settlement_code,settlement_name,street_code,street_name,street_number,building_letter,entrance,apartment_number,floor,epr_address,private_contact_phone,additional_phones,welfare_phone,email,pension,nursing_pension_level,found_in_hazeremim_evacuation_file,wheelchair,foreign_worker,deaf,blind,resides_in_a_nursing_home,lonely,client_type,disability,nursing,mobility,ventilated,special_service,marital_status,number_of_children_under_18,number_of_children,evacuation_status,type_of_evacuation_place,time_in_evacuation_weeks,return_date,updated_residence_place,source_authority,source_settlement,source_of_data_in_population_registry,source_of_data_in_collection_system,source_of_data_in_yachadout_system,source_of_data_in_yachadin_system) " +
                "VALUES (@insert_date, @file_index,@id_number,@id_type,@full_name,@age,@age_group,@gender,@nationality,@number_of_years_in_the_country,@country_of_immigration,@tlm,@settlement_code,@settlement_name,@street_code,@street_name,@street_number,@building_letter,@entrance,@apartment_number,@floor,@epr_address,@private_contact_phone,@additional_phones,@welfare_phone,@email,@pension,@nursing_pension_level,@found_in_hazeremim_evacuation_file,@wheelchair,@foreign_worker,@deaf,@blind,@resides_in_a_nursing_home,@lonely,@client_type,@disability,@nursing,@mobility,@ventilated,@special_service,@marital_status,@number_of_children_under_18,@number_of_children,@evacuation_status,@type_of_evacuation_place,@time_in_evacuation_weeks,@return_date,@updated_residence_place,@source_authority,@source_settlement,@source_of_data_in_population_registry,@source_of_data_in_collection_system,@source_of_data_in_yachadout_system,@source_of_data_in_yachadin_system)";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@insert_date", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                    if (Int32.TryParse(values[0], out int file_index))
                        command.Parameters.AddWithValue("@file_index", file_index);
                    else
                        command.Parameters.AddWithValue("@file_index", 0);

                    command.Parameters.AddWithValue("@id_number", values[1]);
                    command.Parameters.AddWithValue("@id_type", values[2]);
                    command.Parameters.AddWithValue("@full_name", values[3]);
                    if(Int32.TryParse(values[4],out int age))
                        command.Parameters.AddWithValue("@age", age);
                    else
                        command.Parameters.AddWithValue("@age", 0);

                    command.Parameters.AddWithValue("@age_group", values[5]);
                    command.Parameters.AddWithValue("@gender", values[6]);
                    command.Parameters.AddWithValue("@nationality", values[7]);
                    if (Int32.TryParse(values[8], out int number_of_years))
                        command.Parameters.AddWithValue("@number_of_years_in_the_country", number_of_years);
                    else
                        command.Parameters.AddWithValue("@number_of_years_in_the_country", 0);

                    command.Parameters.AddWithValue("@country_of_immigration", values[9]);
                    command.Parameters.AddWithValue("@tlm", values[10]);
                    command.Parameters.AddWithValue("@settlement_code", values[11]);
                    command.Parameters.AddWithValue("@settlement_name", values[12]);
                    command.Parameters.AddWithValue("@street_code", values[13]);
                    command.Parameters.AddWithValue("@street_name", values[14]);
                    command.Parameters.AddWithValue("@street_number", values[15]);
                    command.Parameters.AddWithValue("@building_letter", values[16]);
                    command.Parameters.AddWithValue("@entrance", values[17]);
                    command.Parameters.AddWithValue("@apartment_number", values[18]);
                    if (Int32.TryParse(values[19], out int floor))
                        command.Parameters.AddWithValue("@floor", floor);
                    else
                        command.Parameters.AddWithValue("@floor", 0);

                    command.Parameters.AddWithValue("@epr_address", values[20]);
                    command.Parameters.AddWithValue("@private_contact_phone", values[21]);
                    command.Parameters.AddWithValue("@additional_phones", values[22]);
                    command.Parameters.AddWithValue("@welfare_phone", values[23]);
                    command.Parameters.AddWithValue("@email", values[24]);
                    command.Parameters.AddWithValue("@pension", values[25]);
                    command.Parameters.AddWithValue("@nursing_pension_level", values[26]);
                    command.Parameters.AddWithValue("@found_in_hazeremim_evacuation_file", values[27].ToLower()=="true" ? "True":"False");
                    command.Parameters.AddWithValue("@wheelchair", values[28].ToLower() == "true" ? "True" : "False");
                    command.Parameters.AddWithValue("@foreign_worker", values[29].ToLower() == "true" ? "True" : "False");
                    command.Parameters.AddWithValue("@deaf", values[30].ToLower() == "true" ? "True" : "False");
                    command.Parameters.AddWithValue("@blind", values[31].ToLower() == "true" ? "True" : "False");
                    command.Parameters.AddWithValue("@resides_in_a_nursing_home", values[32].ToLower() == "true" ? "True" : "False");
                    command.Parameters.AddWithValue("@lonely", values[33].ToLower() == "true" ? "True" : "False");
                    command.Parameters.AddWithValue("@client_type", values[34]);
                    command.Parameters.AddWithValue("@disability", values[35]);
                    command.Parameters.AddWithValue("@nursing", values[36].ToLower() == "true" ? "True" : "False");
                    command.Parameters.AddWithValue("@mobility", values[37].ToLower() == "true" ? "True" : "False");
                    command.Parameters.AddWithValue("@ventilated", values[38].ToLower() == "true" ? "True" : "False");
                    command.Parameters.AddWithValue("@special_service", values[39]);
                    command.Parameters.AddWithValue("@marital_status", values[40]);
                    command.Parameters.AddWithValue("@number_of_children_under_18", values[41]);
                    if (Int32.TryParse(values[42], out int number_of_children))
                        command.Parameters.AddWithValue("@number_of_children", number_of_children);
                    else
                        command.Parameters.AddWithValue("@number_of_children", 0);

                    command.Parameters.AddWithValue("@evacuation_status", values[43]);
                    command.Parameters.AddWithValue("@type_of_evacuation_place", values[44]);
                    if (Int32.TryParse(values[45], out int time_in_evacuation_weeks))
                        command.Parameters.AddWithValue("@time_in_evacuation_weeks", time_in_evacuation_weeks);
                    else
                        command.Parameters.AddWithValue("@time_in_evacuation_weeks", 0);
                    
                    if (DateTime.TryParse(values[46], out DateTime returnDate))
                    {
                        command.Parameters.AddWithValue("@return_date", returnDate);
                    }
                    else
                    {
                        // Handle the case where values[46] is not a valid date
                        command.Parameters.AddWithValue("@return_date", DBNull.Value); // or handle appropriately
                    }
                    command.Parameters.AddWithValue("@updated_residence_place", values[47]);
                    command.Parameters.AddWithValue("@source_authority", values[48]);
                    command.Parameters.AddWithValue("@source_settlement", values[49]);
                    command.Parameters.AddWithValue("@source_of_data_in_population_registry", values[50].ToLower() == "true" ? "True" : "False");
                    command.Parameters.AddWithValue("@source_of_data_in_collection_system", values[51].ToLower() == "true" ? "True" : "False");
                    command.Parameters.AddWithValue("@source_of_data_in_yachadout_system", values[52].ToLower() == "true" ? "True" : "False");
                    command.Parameters.AddWithValue("@source_of_data_in_yachadin_system", values[53].ToLower() == "true" ? "True" : "False");


                    connection.Open();
                    command.ExecuteNonQuery();
                    connection.Close();
                }
            }
        }
        
        private void DeleteAllDataFromCurrentDay()
        {
            string query = "Delete from  [dbo].[shibolet_tbl] Where CONVERT(date, insert_date) = CONVERT(date, GETDATE());";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(query, connection))
                {   
                    connection.Open();
                    command.ExecuteNonQuery();
                    connection.Close();
                }
            }
        }
        private void btnSaveDataToExcel_Click(object sender, EventArgs e)
        {
            if (LoadDataFromExcel())
                ReadExcelWithPasswordAndSaveDataToNewExcel();
        }

        private void ReadExcelWithPasswordAndSaveDataToNewExcel()
        {

            FileInfo originalfileInfo = new FileInfo(originalFilePath);
            FileInfo destinationFileInfo = new FileInfo(destinationFilePath);

            // Ensure EPPlus LicenseContext is set
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Read data from the source Excel file
            List<string[]> data = new List<string[]>();
            MessageBox.Show($"password:{password},from file:{originalFilePath}, to file:{destinationFilePath}", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            using (ExcelPackage sourcePackage = new ExcelPackage(originalfileInfo, password))
            {
                ExcelWorksheet sourceWorksheet = sourcePackage.Workbook.Worksheets[0];

                for (int row = 1; row <= sourceWorksheet.Dimension.End.Row; row++)
                {
                    string[] rowData = new string[sourceWorksheet.Dimension.End.Column];
                    for (int col = 1; col <= sourceWorksheet.Dimension.End.Column; col++)
                    {
                        rowData[col - 1] = sourceWorksheet.Cells[row, col].Text;
                    }
                    data.Add(rowData);
                }
            }
            // Check if the file exists
            if (File.Exists(destinationFilePath))
            {
                // Delete the file if it exists
                File.Delete(destinationFilePath);
            }
            // Write data to the new Excel file
            using (ExcelPackage destinationPackage = new ExcelPackage(destinationFileInfo))
            {
                ExcelWorksheet destinationWorksheet = destinationPackage.Workbook.Worksheets.Add("Sheet1");

                for (int row = 0; row < data.Count; row++)
                {
                    for (int col = 0; col < data[row].Length; col++)
                    {
                        if (row>=16 &&  !string.IsNullOrEmpty( data[row][4]))
                        {
                            destinationWorksheet.Cells[row - 16 + 1, col + 1].Value = data[row][col];
                        }
                    }
                }

                // Save the destination package
                destinationPackage.Save();
            }

            MessageBox.Show($"Data successfully read from source and saved to new Excel file.{destinationFilePath}", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}

