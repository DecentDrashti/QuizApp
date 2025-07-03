using System.ComponentModel.DataAnnotations;

namespace QuizApp.Models
{
    public class QuestionLevelModel
    {
        public int? QuestionLevelID {  get; set; }

        [Required(ErrorMessage = "QuestionLevel is required")]
        public String QuestionLevel { get; set; }

        [Required(ErrorMessage = "UserId is required")]
        public int UserID { get; set; }

    }
    public class QuestionLevelDropDownModel
    {
        public int QuestionLevelID { get; set; }
        public string QuestionLevel { get; set; }
    }
}
