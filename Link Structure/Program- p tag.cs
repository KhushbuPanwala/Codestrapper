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
            HtmlDocument document = web.Load("https://nhlink.net/mcq/100-top-smm-multiple-choice-questions-and-answers");


            HtmlNode entry = document.DocumentNode.SelectNodes("//div[@class='entry-content']").First();
            HtmlNode[] questions = entry.SelectNodes("//p").Skip(1).SkipLast(4).ToArray();


            for (int i = 0; i < questions.Count(); i++)
            {
                QuestionAnswer questionAnswer = new QuestionAnswer();



                questionAnswer.QuestionId = (i+1).ToString();
                //questions[i].SelectNodes("strong")[0].InnerText.Split('.')[0].Trim();

                //questionAnswer.QuestionDetails = questions[i].SelectNodes("strong")[0].InnerText.Trim();
                //for (int k = 1; k < questions[i].InnerText.Split('.').Length; k++)
                //{
                //    questionAnswer.QuestionDetails = questionAnswer.QuestionDetails + questions[i].InnerText.Split('.')[k].Trim();

                //}

                questionAnswer.QuestionDetails = questions[i].InnerHtml.Split("<br>")[0].Trim().Replace("<strong>","").Replace("</strong>", "");

                questionAnswer.Options = new List<string>();
                int optLength = questions[i].InnerHtml.Split("<br>").Count();
                for (int k = 1; k < optLength - 1; k++)
                {
                    questionAnswer.Options.Add(questions[i].InnerHtml.Split("<br>")[k].Trim());

                }

                //questionAnswer.Options.Add(questions[i].InnerHtml.Split("<br>")[2].Trim());
                //questionAnswer.Options.Add(questions[i].InnerHtml.Split("<br>")[3].Trim());
                //questionAnswer.Options.Add(questions[i].InnerHtml.Split("<br>")[4].Trim());

                questionAnswer.Answer = questions[i].InnerHtml.Split("<br>")[optLength - 1].Replace("Ans: ", " ").Trim();

                questionAnswer.CompetencyName = "MCQ Questions And Answers";
                questionAnswer.TagName = "SMM";

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
                    singleMultipleQuestionsOptions.IsTrue = item.Split('.')[0].Trim() == questionAnswerList[i].Answer.Trim() ? "TRUE" : "FALSE";

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
                Tuple<string, MemoryStream> fileData = exportToExcelRepository.CreateExcelFileWithMultipleTable(dynamicDictionary, "SMM");

            }
            catch (Exception)
            {
                throw;
            }
            #endregion
        }



    }
}
