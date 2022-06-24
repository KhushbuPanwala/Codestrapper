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

            HtmlWeb web = new HtmlWeb();

            string link = "https://www.javatpoint.com/cloud-computing-mcq";

            HtmlDocument quesionDocument = web.Load(link);

            HtmlNode[] questionsEntry = quesionDocument.DocumentNode.SelectNodes("//p[@class='pq']").ToArray();
            HtmlNode[] optionsEntry = quesionDocument.DocumentNode.SelectNodes("//ol[@class='pointsa']").ToArray();
            HtmlNode[] answersEntry = quesionDocument.DocumentNode.SelectNodes("//div[@class='testanswer']").ToArray();
            //string header = quesionDocument.DocumentNode.SelectNodes("//h2[@class='h2padding']")[0].InnerText.Trim();


            List<QuestionAnswer> questionAnswerList = new List<QuestionAnswer>();
            for (int i = 0; i < questionsEntry.Count(); i++)
            {
                QuestionAnswer questionAnswer = new QuestionAnswer();
                questionAnswer.QuestionId = questionsEntry[i].InnerText.Split(')')[0].Trim();
                //questionAnswer.QuestionDetails = questionsEntry[i].InnerText.Trim();
                int qlength = questionsEntry[i].InnerText.Split(')').Count();
                for (int l = 1; l < qlength; l++)
                {
                    questionAnswer.QuestionDetails = questionAnswer.QuestionDetails + questionsEntry[i].InnerText.Split(')')[l].Trim();
                }

                //options
                List<string> options = new List<string>();

                for (int j = 0; j < optionsEntry[i].SelectNodes("li").Count(); j++)
                {
                    string optionValue = "";
                    if (j == 0)
                    {
                        optionValue = "A. " + optionsEntry[i].SelectNodes("li")[j].InnerText.Trim();
                    }
                    if (j == 1)
                    {
                        optionValue = "B. " + optionsEntry[i].SelectNodes("li")[j].InnerText.Trim();
                    }

                    if (j == 2)
                    {
                        optionValue = "C. " + optionsEntry[i].SelectNodes("li")[j].InnerText.Trim();
                    }

                    if (j == 3)
                    {
                        optionValue = "D. " + optionsEntry[i].SelectNodes("li")[j].InnerText.Trim();
                    }

                    if (j == 4)
                    {
                        optionValue = "E. " + optionsEntry[i].SelectNodes("li")[j].InnerText.Trim();
                    }

                    options.Add(optionValue);
                    //options.Add(optionsEntry[i].SelectNodes("li")[j].InnerText.Trim());
                }
                questionAnswer.Options = options;

                questionAnswer.Answer = answersEntry[i].SelectNodes("p")[0].InnerText.Replace("Answer:", "").Trim();
                //answersEntry[i].SelectNodes("p")[0].SelectSingleNode("strong").InnerText.Trim();
                questionAnswer.CompetencyName = "Multiple-choice Questions";
                //questionAnswer.TagName = header;
                questionAnswer.QuestionType = "Single";
                questionAnswer.DifficultyLevel = "Medium";

                questionAnswerList.Add(questionAnswer);

            }




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
                    singleMultipleQuestionsOptions.OptionDetail = item.Split(". ")[1];
                    //.Split('.')[1];

                    singleMultipleQuestionsOptions.IsTrue = item.Split(".")[0].ToString() == questionAnswerList[i].Answer.ToString() ? "TRUE" : "FALSE";

                    //singleMultipleQuestionsOptions.IsTrue = singleMultipleQuestionsOptions.OptionDetail.Split(")")[0].ToString() == questionAnswerList[i].Answer ? "TRUE" : "FALSE";

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
                Tuple<string, MemoryStream> fileData = exportToExcelRepository.CreateExcelFileWithMultipleTable(dynamicDictionary, "Output");

            }
            catch (Exception)
            {
                throw;
            }
            #endregion
        }



    }
}
