using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace QuizApp.Models
{
    public class UserModel
    {
        public int? UserID { get; set; }

        [Required(ErrorMessage = "UserName is required")]
        public String UserName { get; set; }

        [Required(ErrorMessage ="Password is Required")]
        [PasswordPropertyText]
        public String Password { get; set; }

        [Required(ErrorMessage = "EmailID is required")]
        [EmailAddress]
        public string Email { get; set; }

        [MaxLength(10, ErrorMessage = "Mobileno is required of 10 digit")]
        public String Mobile { get; set; }

        public bool IsActive { get; set; } = true;


    }
    public class UserDropDownModel
    {
        public int UserID { get; set; }
        public string UserName { get; set; }
    }

    public class UserLoginModel
    {
        [Required]
        public string UserName { get; set; }
        [Required]
        public string Password { get; set; }
    }
}
