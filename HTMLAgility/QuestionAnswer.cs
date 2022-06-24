using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace HTMLAgility
{
    class QuestionAnswer
    {
        public string QuestionId { get; set; }
        public string QuestionType { get; set; }
        public string DifficultyLevel { get; set; }
        public string QuestionDetails { get; set; }

        public List<string> OptionsIds { get; set; }
        public List<string> Options { get; set; }

        public string TagName { get; set; }
        public string CompetencyName { get; set; }
        public string Answer { get; set; }

    }



    public class Questions
    {
        [DisplayName("Question Id(Required")]
        public string QuestionId { get; set; }

        [DisplayName("Question Type - Single/Multiple(Required")]
        public string QuestionType { get; set; }

        [DisplayName("Difficulty level(Required")]
        public string DifficultyLevel { get; set; }
        [DisplayName("Question Details(Required")]
        public string QuestionDetails { get; set; }

        [DisplayName("Basic or Premium(Required")]
        public string BasicOrPremium { get; set; }
    }


    class QuestionTags
    {
        [DisplayName("Question Id(Required)")]
        public string QuestionId { get; set; }
        [DisplayName("Tag Name(Required)")]
        public string TagName { get; set; }
        [DisplayName("Competency Name(Required)")]
        public string CompetencyName { get; set; }

    }



    class SingleMultipleQuestionsOptions
    {

        [DisplayName("Question Id(Only of Single/Multiple Question Types)")]
        public string QuestionId { get; set; }

        [DisplayName("Option Detail(Required)")]
        public string OptionDetail { get; set; }

        [DisplayName("Is True(Required)")]
        public string IsTrue { get; set; }

    }

    public class Instructions
    {
        [DisplayName("Title")]
        public string Title { get; set; }
    }


}
