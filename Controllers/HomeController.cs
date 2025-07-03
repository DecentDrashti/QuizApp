using System.Data.SqlClient;
using System.Data;
using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using QuizApp.Models;

namespace QuizApp.Controllers
{
    [CheckAccess]
    public class HomeController : Controller
    {
        private readonly IConfiguration _configuration;

        public HomeController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public IActionResult Index()
        {
            string connectionStr = _configuration.GetConnectionString("ConnectionString");

            HomeModel model = new HomeModel();

            using (SqlConnection conn = new SqlConnection(connectionStr))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand("PR_Count", conn);
                cmd.CommandType = CommandType.StoredProcedure;

                SqlDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    model.QuestionCount = Convert.ToInt32(reader["QuestionCount"]);
                    model.QuestionLevelCount = Convert.ToInt32(reader["QuestionLevelCount"]);
                    model.QuizCount = Convert.ToInt32(reader["QuizCount"]);
                    model.UserCount = Convert.ToInt32(reader["UserCount"]);
                }
            }

            return View("Index", model);
        }
    }
}
