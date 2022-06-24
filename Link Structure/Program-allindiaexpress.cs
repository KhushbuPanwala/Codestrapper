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

            List<QuestionAnswer> questionAnswerList = new List<QuestionAnswer>();
            string tagName = "";
            for (int k = 1; k <= 5; k++)
            {

                string link = "http://www.allindiaexams.in/engineering/cse/javascript-mcq/invocation-performance-navigation/" + k;
                HtmlDocument quesionDocument = web.Load(link);

                HtmlNode[] entry1 = quesionDocument.DocumentNode.SelectNodes("//div[@class='qa_list']").ToArray();

                string header = quesionDocument.DocumentNode.SelectNodes("//div[@class='int_content']")[0].SelectNodes("h1")[0].InnerText;

                List<HtmlNode> entry2 = new List<HtmlNode>();
                #region questions
                for (int i = 0; i < entry1.Count(); i++)
                {
                    if (entry1[i].InnerText.Trim() != "")
                    {
                        entry2.Add(entry1[i]);
                    }
                }


                for (int i = 0; i < entry2.Count(); i++)
                {
                    QuestionAnswer questionAnswer = new QuestionAnswer();
                    questionAnswer.QuestionId = entry2[i].SelectSingleNode(".//span[@class='sno']").InnerText.Split(".")[0];
                    questionAnswer.QuestionDetails = entry2[i].SelectSingleNode(".//span[@class='sno']/following-sibling::text()").InnerText.Trim();
                    //options
                    List<string> options = new List<string>();

                    var optionsDataCount = entry2[i].SelectNodes("//ul[@class='options_list clearfix']")[i].InnerText.Trim().Split("\r\n\t\t\t\t\t").Count();
                    var optionsData = entry2[i].SelectNodes("//ul[@class='options_list clearfix']")[i];
                    for (int j = 0; j < optionsDataCount; j++)
                    {
                        options.Add(optionsData.InnerText.Trim().Split("\r\n\t\t\t\t\t")[j].Trim());
                    }
                    questionAnswer.Options = options;
                    questionAnswer.Answer = entry2[i].SelectNodes("section")[0].SelectNodes("div")[0].SelectNodes("p")[0].InnerHtml.Trim();
                    //questionAnswer.CompetencyName = header.Split('-')[1];
                    //questionAnswer.TagName = header.Split('-')[0];
                    questionAnswer.TagName = header.Replace(" Multiple Choice Questions and Answers", "");
                    tagName = questionAnswer.TagName;
                    questionAnswer.CompetencyName = "Multiple Choice Questions and Answers";
                    questionAnswer.QuestionType = "Single";
                    questionAnswer.DifficultyLevel = "Medium";

                    questionAnswerList.Add(questionAnswer);
                }
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
                    singleMultipleQuestionsOptions.OptionDetail = item;
                    //.Split('.')[1];

                    singleMultipleQuestionsOptions.IsTrue = item.Split(".")[0].ToString() == questionAnswerList[i].Answer.Split(' ')[2].ToString() ? "TRUE" : "FALSE";

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
                Tuple<string, MemoryStream> fileData = exportToExcelRepository.CreateExcelFileWithMultipleTable(dynamicDictionary, tagName);

            }
            catch (Exception)
            {
                throw;
            }
            #endregion
        }



    }
}
