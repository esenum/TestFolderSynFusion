using Syncfusion.Pdf.Parsing;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;

namespace PDFFormFillingSample
{
    class Program
    {
        static void Main(string[] args)
        {
            //Connection string 
            string dbFile = System.IO.Path.GetFullPath("../../Data/") + "Database2.mdb";
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbFile;
            //SQL Query
            string strSQL = "SELECT * From Table1";

            //Create a connection  
            OleDbConnection connection = new OleDbConnection(connectionString);
            //Create a command and set its connection  
            OleDbCommand command = new OleDbCommand(strSQL, connection);
            //Open the connection and execute the select command
            connection.Open();
            //Execute command  
            OleDbDataReader reader = command.ExecuteReader();

            //Load the PDF document 
            PdfLoadedDocument loadedDocument = new PdfLoadedDocument("../../Data/FormTemplate.pdf");
            //Get the loaded form
            PdfLoadedForm loadedForm = loadedDocument.Form;

            while (reader.Read())
            {
                //Get the values from database
                string name = reader["Name"].ToString();
                string phone = reader["Phone"].ToString();
                string email = reader["Email Address"].ToString();

                //Assign values to fields
                (loadedForm.Fields["Name"] as PdfLoadedTextBoxField).Text = name;
                (loadedForm.Fields["Phone"] as PdfLoadedTextBoxField).Text = phone;
                (loadedForm.Fields["Email address"] as PdfLoadedTextBoxField).Text = email;
            }

            //SQL Query for another table
            strSQL = "select * from Table2";
            //Create a connection  
            connection = new OleDbConnection(connectionString);
            //Create a command and set its connection  
            command = new OleDbCommand(strSQL, connection);
            //Open the connection and execute the select command.  
            connection.Open();
            //Execute command  
            reader = command.ExecuteReader();

            while (reader.Read())
            {
                //Get values from database
                string qualification = reader["Qualification"].ToString();
                bool isSelected = Convert.ToBoolean(reader["IsSelected"]);

                //Check the checkbox
                PdfLoadedCheckBoxField loadedCheckBoxField = loadedForm.Fields[qualification] as PdfLoadedCheckBoxField;
                loadedCheckBoxField.Checked = isSelected;
            }
            //Save the document
            loadedDocument.Save("Form.pdf");

            //Close the document 
            loadedDocument.Close(true);

            //This will open the PDF file so, the result will be seen in default PDF Viewer 
            Process.Start("Form.pdf");
        }
    }
}
