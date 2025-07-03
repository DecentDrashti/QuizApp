using System.Data.SqlClient;
using System.Data;
using Microsoft.AspNetCore.Mvc;
using QuizApp.Models;

namespace QuizApp.Controllers
{
    public class RegisterController : Controller
    {
        private IConfiguration Configuration;

        public RegisterController(IConfiguration _configuration)
        {
            Configuration = _configuration;
        }
        
        public IActionResult UserRegister(UserModel UserModel)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    string connectionString = this.Configuration.GetConnectionString("ConnectionString");
                    SqlConnection sqlConnection = new SqlConnection(connectionString);
                    sqlConnection.Open();
                    SqlCommand sqlCommand = sqlConnection.CreateCommand();
                    sqlCommand.CommandType = System.Data.CommandType.StoredProcedure;
                    sqlCommand.CommandText = "PR_User_Insert";
                    sqlCommand.Parameters.Add("@UserName", SqlDbType.VarChar).Value = UserModel.UserName;
                    sqlCommand.Parameters.Add("@Password", SqlDbType.VarChar).Value = UserModel.Password;
                    sqlCommand.Parameters.Add("@Mobile", SqlDbType.VarChar).Value = UserModel.Mobile;
                    sqlCommand.Parameters.Add("@Email", SqlDbType.VarChar).Value = UserModel.Email;
                    sqlCommand.Parameters.Add("@IsActive", SqlDbType.VarChar).Value = UserModel.IsActive;
                    sqlCommand.ExecuteNonQuery();
                    return RedirectToAction("Login", "Login");
                }
            }
            catch (Exception e)
            {
                TempData["ErrorMessage"] = e.Message;
                return RedirectToAction("Register");
            }
            return RedirectToAction("Register");
        }
        public IActionResult Register()
        {
            return View();
        }
    }
}
