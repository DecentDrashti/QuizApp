using System.ComponentModel.DataAnnotations;

namespace QuizApp.Models
{
    public class QuizWiseQuestionModel
    {
        public int? QuizWiseQuestionID { get; set; }

        [Required(ErrorMessage = "QuizId is required")]
        public int QuizID { get; set; }

        [Required(ErrorMessage = "QuestionID is required")]
        public int QuestionID { get; set; }

        [Required(ErrorMessage = "UserID is required")]
        public int UserID { get; set; }

    }
}
