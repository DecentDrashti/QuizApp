using System.Data.SqlClient;
using System.Data;
using Microsoft.AspNetCore.Mvc;
using QuizApp.Models;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;

namespace QuizApp.Controllers
{
    [CheckAccess]
    public class QuestionController : Controller
    {
        private IConfiguration configuration;

        public QuestionController(IConfiguration _configuration)
        {
            configuration = _configuration;
        }

        public IActionResult QuestionList()
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Question_SelectAll";
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            return View(table);
        }
        public IActionResult QuestionAddEdit(QuestionModel model)
        {
            if (ModelState.IsValid)
            {
                string connectionString = this.configuration.GetConnectionString("ConnectionString");
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;

                if (model.QuestionID == null)
                {
                    command.CommandText = "PR_Question_Insert";
                }
                else
                {
                    command.CommandText = "PR_Question_UpdateByPK";
                    command.Parameters.Add("@QuestionID", SqlDbType.Int).Value = model.QuestionID;
                }
                command.Parameters.Add("@QuestionText", SqlDbType.VarChar).Value = model.QuestionText;
                command.Parameters.Add("@QuestionLevelID", SqlDbType.Int).Value = model.QuestionLevelID;
                command.Parameters.Add("@OptionA", SqlDbType.VarChar).Value = model.OptionA;
                command.Parameters.Add("@OptionB", SqlDbType.VarChar).Value = model.OptionB;
                command.Parameters.Add("@OptionC", SqlDbType.VarChar).Value = model.OptionC;
                command.Parameters.Add("@OptionD", SqlDbType.VarChar).Value = model.OptionD;
                command.Parameters.Add("@CorrectOption", SqlDbType.VarChar).Value = model.CorrectOption;
                command.Parameters.Add("@QuestionMarks", SqlDbType.Int).Value = model.QuestionMarks;
                command.Parameters.Add("@IsActive", SqlDbType.Bit).Value = model.IsActive;
                command.Parameters.Add("@UserID", SqlDbType.Int).Value = model.UserID;
                command.ExecuteNonQuery();
                return RedirectToAction("QuestionList");
            }

            return View("QuestionForm", model);
        }
        public IActionResult QuestionForm(int? QuestionID)
        {
            QuestionLevelDropDown();
            UserDropDown();
            if (QuestionID == null)
            {
                return View();
            }

            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_Question_SelectById";
            command.Parameters.AddWithValue("@QuestionID", QuestionID);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            QuestionModel model = new QuestionModel();

            foreach (DataRow dataRow in table.Rows)
            {
                model.QuestionID = Convert.ToInt32(dataRow["QuestionID"]);
                model.QuestionText = dataRow["QuestionText"].ToString();
                model.QuestionLevelID = Convert.ToInt32(dataRow["QuestionLevelID"]);
                model.OptionA = dataRow["OptionA"].ToString();
                model.OptionB = dataRow["OptionB"].ToString();
                model.OptionC = dataRow["OptionC"].ToString();
                model.OptionD = dataRow["OptionD"].ToString();
                model.CorrectOption = dataRow["CorrectOption"].ToString();
                model.QuestionMarks = Convert.ToInt32(dataRow["QuestionMarks"]);
                model.IsActive = Convert.ToBoolean(dataRow["IsActive"]);
                model.UserID = Convert.ToInt32(dataRow["UserID"]);
            }
            return View(model);
        }

        public void UserDropDown()
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command2 = connection.CreateCommand();
            command2.CommandType = System.Data.CommandType.StoredProcedure;
            command2.CommandText = "PR_User_DropDown";
            SqlDataReader reader2 = command2.ExecuteReader();
            DataTable dataTable2 = new DataTable();
            dataTable2.Load(reader2);
            List<UserDropDownModel> UserList = new List<UserDropDownModel>();
            foreach (DataRow data in dataTable2.Rows)
            {
                UserDropDownModel model = new UserDropDownModel();
                model.UserID = Convert.ToInt32(data["UserID"]);
                model.UserName = data["UserName"].ToString();
                UserList.Add(model);
            }
            ViewBag.UserList = UserList;
            //.Count > 0 ? QuizList : new List<QuizDropDownModel>();
        }

        public void QuestionLevelDropDown()
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command2 = connection.CreateCommand();
            command2.CommandType = System.Data.CommandType.StoredProcedure;
            command2.CommandText = "PR_QuestionLevel_DropDown";
            SqlDataReader reader2 = command2.ExecuteReader();
            DataTable dataTable2 = new DataTable();
            dataTable2.Load(reader2);
            List<QuestionLevelDropDownModel> QuestionLevelList = new List<QuestionLevelDropDownModel>();
            foreach (DataRow data in dataTable2.Rows)
            {
                QuestionLevelDropDownModel model = new QuestionLevelDropDownModel();
                model.QuestionLevelID = Convert.ToInt32(data["QuestionLevelID"]);
                model.QuestionLevel = data["QuestionLevel"].ToString();
                QuestionLevelList.Add(model);
            }
            ViewBag.QuestionLevelList = QuestionLevelList;
            //.Count > 0 ? QuizList : new List<QuizDropDownModel>();
        }

        public IActionResult QuestionDropDownAddEdit(QuestionModel model)
        {
            if (ModelState.IsValid)
            {
                return RedirectToAction("QuestionList");
            }
            QuestionLevelDropDown();
            UserDropDown();
            return View("QuestionForm", model);
        }

        public IActionResult QuestionExportToExcel()
        {
            string connectionString = configuration.GetConnectionString("ConnectionString");
            SqlConnection sqlConnection = new SqlConnection(connectionString);
            sqlConnection.Open();

            SqlCommand sqlCommand = sqlConnection.CreateCommand();
            sqlCommand.CommandType = System.Data.CommandType.StoredProcedure;
            sqlCommand.CommandText = "PR_Question_SelectAll";
            //sqlCommand.Parameters.Add("@UserID", SqlDbType.Int).Value = CommonVariable.UserID();

            SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
            DataTable data = new DataTable();
            data.Load(sqlDataReader);

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("DataSheet");

                // Add headers
                worksheet.Cells[1, 1].Value = "QuestionID";
                worksheet.Cells[1, 2].Value = "QuestionText";
                worksheet.Cells[1, 3].Value = "QuestionLevelID";
                worksheet.Cells[1, 4].Value = "OptionA";
                worksheet.Cells[1, 5].Value = "OptionB";
                worksheet.Cells[1, 6].Value = "OptionC";
                worksheet.Cells[1, 7].Value = "OptionD";
                worksheet.Cells[1, 8].Value = "CorrectOption";
                worksheet.Cells[1, 9].Value = "QuestionMarks";
                worksheet.Cells[1, 10].Value = "IsActive";
                worksheet.Cells[1, 11].Value = "UserID";
                worksheet.Cells[1, 12].Value = "Created";
                worksheet.Cells[1, 13].Value = "Modified";

                // Add data
                int row = 2;
                foreach (DataRow item in data.Rows)
                {
                    
                    worksheet.Cells[row, 1].Value = item["QuestionID"];
                    worksheet.Cells[row, 2].Value = item["QuestionText"];
                    worksheet.Cells[row, 3].Value = item["QuestionLevelID"];
                    worksheet.Cells[row, 4].Value = item["OptionA"];
                    worksheet.Cells[row, 5].Value = item["OptionB"];
                    worksheet.Cells[row, 6].Value = item["OptionC"];
                    worksheet.Cells[row, 7].Value = item["OptionD"];
                    worksheet.Cells[row, 8].Value = item["CorrectOption"];
                    worksheet.Cells[row, 9].Value = item["QuestionMarks"];
                    worksheet.Cells[row, 10].Value = item["IsActive"];
                    worksheet.Cells[row, 11].Value = item["UserID"];
                    worksheet.Cells[row, 12].Value = item["Created"];
                    worksheet.Cells[row, 13].Value = item["Modified"];
                    row++;
                }

                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

                string excelName = $"Data-{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
            }
        }

        public IActionResult QuestionDelete(int QuestionID)
        {
            Console.WriteLine("Delete method called for QuestionID: " + QuestionID);

            try
            {
                string connectionString = configuration.GetConnectionString("ConnectionString");
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = connection.CreateCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "PR_Question_DeleteByPK";
                    command.Parameters.Add("@QuestionID", SqlDbType.Int).Value = QuestionID;
                    command.ExecuteNonQuery();
                }

                TempData["SuccessMessage"] = "Question deleted successfully.";
                return RedirectToAction("QuestionList");
            }
            catch (Exception ex)
            {
                TempData["ErrorMessage"] = "An error occurred while deleting the Question: " + ex.Message;
                return RedirectToAction("QuestionList");
            }
        }
    }
}
