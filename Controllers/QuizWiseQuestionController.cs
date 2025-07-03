using System.Data.SqlClient;
using System.Data;
using Microsoft.AspNetCore.Mvc;
using QuizApp.Models;
using QuizManagement.Models;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;

namespace QuizApp.Controllers
{
    [CheckAccess]
    public class QuizWiseQuestionController : Controller
    {
        private IConfiguration configuration;

        public QuizWiseQuestionController(IConfiguration _configuration)
        {
            configuration = _configuration;
        }

        public IActionResult QuizWiseQuestionList()
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_QuizWiseQuestion_SelectAll";
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            return View(table);
        }
        public IActionResult QuizWiseQuestionAddEdit(QuizWiseQuestionModel model)
        {
            if (ModelState.IsValid)
            {
                string connectionString = this.configuration.GetConnectionString("ConnectionString");
                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand command = connection.CreateCommand();
                command.CommandType = CommandType.StoredProcedure;

                if (model.QuizWiseQuestionID == null)
                {
                    command.CommandText = "PR_QuizWiseQuestion_Insert";
                }
                else
                {
                    command.CommandText = "PR_QuizWiseQuestion_UpdateByPK";
                    command.Parameters.Add("@QuizWiseQuestionID", SqlDbType.Int).Value = model.QuizWiseQuestionID;
                }
                command.Parameters.Add("@QuizID", SqlDbType.VarChar).Value = model.QuizID;
                command.Parameters.Add("@QuestionID", SqlDbType.VarChar).Value = model.QuestionID;
                command.Parameters.Add("@UserID", SqlDbType.VarChar).Value = model.UserID;
                command.ExecuteNonQuery();
                return RedirectToAction("QuizWiseQuestionList");
            }
            return View("QuizWiseQuestionForm", model);
        }
        public IActionResult QuizWiseQuestionForm(int? QuizWiseQuestionID)
        {
            QuizDropDown();
            QuestionDropDown();
            UserDropDown();
            if (QuizWiseQuestionID == null)
            {
                return View();
            }

            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command = connection.CreateCommand();
            command.CommandType = CommandType.StoredProcedure;
            command.CommandText = "PR_QuizWiseQuestion_SelectById";
            command.Parameters.AddWithValue("@QuizWiseQuestionID", QuizWiseQuestionID);
            SqlDataReader reader = command.ExecuteReader();
            DataTable table = new DataTable();
            table.Load(reader);
            QuizWiseQuestionModel model = new QuizWiseQuestionModel();

            foreach (DataRow dataRow in table.Rows)
            {
                model.QuizWiseQuestionID = Convert.ToInt32(dataRow["QuizWiseQuestionID"]);
                model.QuizID = Convert.ToInt32(dataRow["QuizID"]);
                model.QuestionID = Convert.ToInt32(dataRow["QuestionID"]);
                model.UserID = Convert.ToInt32(dataRow["UserID"]);
            }
            return View(model);
        }

        public void QuizDropDown()
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command2 = connection.CreateCommand();
            command2.CommandType = System.Data.CommandType.StoredProcedure;
            command2.CommandText = "PR_Quiz_DropDown";
            SqlDataReader reader2 = command2.ExecuteReader();
            DataTable dataTable2 = new DataTable();
            dataTable2.Load(reader2);
            List<QuizDropDownModel> QuizList = new List<QuizDropDownModel>();
            foreach (DataRow data in dataTable2.Rows)
            {
                QuizDropDownModel model = new QuizDropDownModel();
                model.QuizID = Convert.ToInt32(data["QuizID"]);
                model.QuizName = data["QuizName"].ToString();
                QuizList.Add(model);
            }
            ViewBag.QuizList = QuizList;
                //.Count > 0 ? QuizList : new List<QuizDropDownModel>();
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

        public void QuestionDropDown()
        {
            string connectionString = this.configuration.GetConnectionString("ConnectionString");
            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand command2 = connection.CreateCommand();
            command2.CommandType = System.Data.CommandType.StoredProcedure;
            command2.CommandText = "PR_Question_DropDown";
            SqlDataReader reader2 = command2.ExecuteReader();
            DataTable dataTable2 = new DataTable();
            dataTable2.Load(reader2);
            List<QuestionDropDownModel> QuestionList = new List<QuestionDropDownModel>();
            foreach (DataRow data in dataTable2.Rows)
            {
                QuestionDropDownModel model = new QuestionDropDownModel();
                model.QuestionID = Convert.ToInt32(data["QuestionID"]);
                model.QuestionText = data["QuestionText"].ToString();
                QuestionList.Add(model);
            }
            ViewBag.QuestionList = QuestionList;
            //.Count > 0 ? QuizList : new List<QuizDropDownModel>();
        }

        public IActionResult QuizWiseQuestionsDropDownAddEdit(QuizWiseQuestionModel model)
        {
            if (ModelState.IsValid)
            {
                return RedirectToAction("QuizWiseQuestionList");
            }

            QuizDropDown();
            QuestionDropDown();
            UserDropDown();
            return View("QuizWiseQuestionForm", model);
        }

        public IActionResult QuizWiseQuestionExportToExcel()
        {
            string connectionString = configuration.GetConnectionString("ConnectionString");
            SqlConnection sqlConnection = new SqlConnection(connectionString);
            sqlConnection.Open();

            SqlCommand sqlCommand = sqlConnection.CreateCommand();
            sqlCommand.CommandType = System.Data.CommandType.StoredProcedure;
            sqlCommand.CommandText = "PR_QuizWiseQuestion_SelectAll";
            //sqlCommand.Parameters.Add("@UserID", SqlDbType.Int).Value = CommonVariable.UserID();

            SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
            DataTable data = new DataTable();
            data.Load(sqlDataReader);

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("DataSheet");

                // Add headers
                worksheet.Cells[1, 1].Value = "QuizWiseQuestionID";
                worksheet.Cells[1, 2].Value = "QuizID";
                worksheet.Cells[1, 3].Value = "QuestionID";
                worksheet.Cells[1, 4].Value = "UserID";
                worksheet.Cells[1, 5].Value = "Created";
                worksheet.Cells[1, 6].Value = "Modified";

                // Add data
                int row = 2;
                foreach (DataRow item in data.Rows)
                {
                    worksheet.Cells[row, 1].Value = item["QuizWiseQuestionID"];
                    worksheet.Cells[row, 2].Value = item["QuizID"];
                    worksheet.Cells[row, 3].Value = item["QuestionID"];
                    worksheet.Cells[row, 4].Value = item["UserID"];
                    worksheet.Cells[row, 5].Value = item["Created"];
                    worksheet.Cells[row, 6].Value = item["Modified"];
                    row++;
                }

                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

                string excelName = $"Data-{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
            }
        }

        public IActionResult QuizWiseQuestionDelete(int QuizWiseQuestionID)
        {
            try
            {
                string connectionString = configuration.GetConnectionString("ConnectionString");
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlCommand command = connection.CreateCommand();
                    command.CommandType = CommandType.StoredProcedure;
                    command.CommandText = "PR_QuizWiseQuestion_DeleteByPK";
                    command.Parameters.Add("@QuizWiseQuestionID", SqlDbType.Int).Value = QuizWiseQuestionID;
                    command.ExecuteNonQuery();
                }

                TempData["SuccessMessage"] = "QuizWiseQuestion deleted successfully.";
                return RedirectToAction("QuizWiseQuestionList");
            }
            catch (Exception ex)
            {
                TempData["ErrorMessage"] = "An error occurred while deleting the QuizWiseQuestion: " + ex.Message;
                return RedirectToAction("QuizWiseQuestionList");
            }
        }
    }
}