using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace HTMLAgility
{
    class Program
    {
        static void Main(string[] args)
        {

            List<QuestionAnswer> questionAnswerList = new List<QuestionAnswer>();

            #region set question and options  
            HtmlWeb web = new HtmlWeb();
            HtmlDocument document = web.Load("https://www.nngroup.com/articles/ux-quiz/");


            HtmlNode entry = document.DocumentNode.SelectNodes("//div[@class='publication-container']//section[@class='article-body']").First();
            HtmlNode[] nodes = entry.SelectNodes("//ol//li").ToArray();


            HtmlNode[] questions = entry.SelectNodes("//ol//li//strong").ToArray();

            HtmlNode[] options = entry.SelectNodes("//ol//li//ol//li").ToArray();
            //HtmlNode[] answers = entry.SelectNodes("//ol//p").ToArray();
            //List<string> answerList = new List<string>();
            //for (int i = 0; i < answers.Count(); i++)
            //{
            //    if (!string.IsNullOrEmpty(answers[i].InnerText))
            //    {
            //        answerList.Add(answers[i].InnerText);
            //    }
            //}

            int start = 0;
            int end = 0;

            for (int i = 0; i < questions.Count(); i++)
            {
                QuestionAnswer questionAnswer = new QuestionAnswer();

                //questionAnswer.QuestionDetails = nodes[i].SelectNodes("//strong").First().InnerText.Trim();
                questionAnswer.QuestionId = (i + 1).ToString();
                questionAnswer.QuestionDetails = questions[i].InnerText.Trim();

                questionAnswer.Options = new List<string>();
                //HtmlNode[] options = nodes[i].SelectNodes("ol//li").ToArray();

                if (end == 0)
                {
                    start = i;
                    end = end + 4;
                }
                else
                {
                    start = end;
                    end = end + 4;

                }
                for (int j = start; j < end; j++)
                {
                    questionAnswer.Options.Add(options[j].InnerText.Trim());

                }


                //questionAnswer.Answer = answerList[i].Trim();

                questionAnswer.CompetencyName = "MCQ Questions And Answers";
                questionAnswer.TagName = "User-Experience";

                questionAnswer.QuestionType = "Single";
                questionAnswer.DifficultyLevel = "Medium";

                questionAnswerList.Add(questionAnswer);
            }

            #endregion
            #region Bind data for excel
            Program generateExcel = new Program();
            List<Questions> questionDetails = new List<Questions>();
            for (int i = 0; i < questionAnswerList.Count; i++)
            {
                Questions questionSheet = new Questions();
                questionSheet.QuestionId = questionAnswerList[i].QuestionId;
                questionSheet.QuestionType = questionAnswerList[i].QuestionType;
                questionSheet.DifficultyLevel = questionAnswerList[i].DifficultyLevel;
                questionSheet.QuestionDetails = questionAnswerList[i].QuestionDetails;
                questionSheet.BasicOrPremium = string.Empty;
                questionDetails.Add(questionSheet);
            }



            List<QuestionTags> questionTagDetails = new List<QuestionTags>();
            for (int i = 0; i < questionAnswerList.Count; i++)
            {
                QuestionTags questionTags = new QuestionTags();
                questionTags.QuestionId = questionAnswerList[i].QuestionId;
                questionTags.TagName = questionAnswerList[i].TagName;
                questionTags.CompetencyName = questionAnswerList[i].CompetencyName;
                questionTagDetails.Add(questionTags);
            }

            List<SingleMultipleQuestionsOptions> singleMultipleQuestionsOptionDetails = new List<SingleMultipleQuestionsOptions>();
            for (int i = 0; i < questionAnswerList.Count; i++)
            {
                SingleMultipleQuestionsOptions singleMultipleQuestionsOptions = new SingleMultipleQuestionsOptions();
                foreach (var item in questionAnswerList[i].Options)
                {
                    singleMultipleQuestionsOptions = new SingleMultipleQuestionsOptions();

                    singleMultipleQuestionsOptions.QuestionId = questionAnswerList[i].QuestionId;
                    singleMultipleQuestionsOptions.OptionDetail = item.Trim();
                    singleMultipleQuestionsOptions.IsTrue = "FALSE";
                    //item.Trim() == questionAnswerList[i].Answer.Trim() ? "TRUE" : "FALSE";

                    singleMultipleQuestionsOptionDetails.Add(singleMultipleQuestionsOptions);
                }

            }


            #endregion

            #region Create Excel
            //create dynamic directory
            dynamic dynamicDictionary = new DynamicDictionary<string, dynamic>();
            dynamicDictionary.Add("Question", questionDetails);
            dynamicDictionary.Add("Question Tags", questionTagDetails);
            dynamicDictionary.Add("SingleMultipleQuestionsOptions", singleMultipleQuestionsOptionDetails);

            try
            {
                ExportToExcelRepository exportToExcelRepository = new ExportToExcelRepository();
                Tuple<string, MemoryStream> fileData = exportToExcelRepository.CreateExcelFileWithMultipleTable(dynamicDictionary, "User-Experience");

            }
            catch (Exception)
            {
                throw;
            }
            #endregion
        }

    }
}
