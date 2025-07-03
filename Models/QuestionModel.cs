using System.ComponentModel.DataAnnotations;

namespace QuizApp.Models
{
    public class QuestionModel
    {
        public int? QuestionID { get; set; }

        [Required(ErrorMessage = "QuestionText is required")]
        public String QuestionText { get; set; }

        [Required(ErrorMessage = "QuestionLevelID is required")]
        public int QuestionLevelID { get; set; }

        [Required(ErrorMessage = "OptionA is required")]
        public String OptionA { get; set; }

        [Required(ErrorMessage = "OptionB is required")]
        public String OptionB { get; set; }

        [Required(ErrorMessage = "OptionC is required")]
        public String OptionC { get; set; }

        [Required(ErrorMessage = "OptionD is required")]
        public String OptionD { get; set; }

        [Required(ErrorMessage = "CorrectOption is required")]
        public String CorrectOption { get; set; }

        [Required(ErrorMessage = "QuestionMarks is required")]
        public int QuestionMarks { get; set; }

        public bool IsActive { get; set; } = true;

        [Required(ErrorMessage = "UserId is required")]
        public int UserID { get; set; }

    }
    public class QuestionDropDownModel
    {
        public int QuestionID { get; set; }
        public string QuestionText { get; set; }
    }
}
