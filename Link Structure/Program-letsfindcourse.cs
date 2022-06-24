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
            HtmlDocument document = web.Load("https://letsfindcourse.com/technical-questions/digital-marketing-mcq/digital-marketing-mcq");


            HtmlNode[] links = document.DocumentNode.SelectNodes("//div[@class='col-3 col-lg-3 col-md-3 mcqtopic']//ul//li//a").ToArray();

            HtmlNode[] headerData = document.DocumentNode.SelectNodes("//div[@class='col-3 col-lg-3 col-md-3 mcqtopic']//ul//li").Skip(1).ToArray();
            List<string> linkArray = new List<string>();
            foreach (HtmlNode item in links)
            {
                HtmlAttribute att = item.Attributes["href"];
                if (att.Value.Contains("https://"))
                {
                    linkArray.Add(att.Value);
                }
                else
                {

                    linkArray.Add("https://letsfindcourse.com/technical-questions/digital-marketing-mcq/" + att.Value);
                }
            }
            List<string> headerArray = new List<string>();
            foreach (HtmlNode item in headerData)
            {

                Console.WriteLine(item.InnerText.Trim());
                headerArray.Add(item.InnerText.Trim());
            }

            for (int k = 0; k < linkArray.Count; k++)
            {
                List<QuestionAnswer> questionAnswerList = new List<QuestionAnswer>();

                #region set question and options  

                HtmlDocument quesionDocument = web.Load(linkArray[k]);

                HtmlNode entry = quesionDocument.DocumentNode.SelectNodes("//div[@class='content']").First();
                HtmlNode[] questions = entry.SelectNodes("//p[@class='mcq']").ToArray();
                HtmlNode[] options = entry.SelectNodes("//p[@class='options']").ToArray();

                HtmlNode[] answers = entry.SelectNodes("//div[@class='showanswer']").ToArray();
                //HtmlNode header = entry.SelectNodes("//h1").First();

                for (int i = 0; i < questions.Count(); i++)
                {
                    QuestionAnswer questionAnswer = new QuestionAnswer();

                    questionAnswer.QuestionId = questions[i].InnerText.Split("&nbsp;")[0].Trim();
                    questionAnswer.QuestionDetails = questions[i].InnerText.Split("&nbsp;")[1].Trim();

                    questionAnswer.Options = new List<string>();

                    List<string> optionsDetail = options[i].InnerText.Split("\r\n").ToList();
                    for (int j = 0; j < optionsDetail.Count; j++)
                    {
                        if (!string.IsNullOrEmpty(optionsDetail[j].Trim()))
                        {

                            questionAnswer.Options.Add(optionsDetail[j].Trim());
                        }

                    }

                    questionAnswer.Answer = answers[i].InnerHtml.Split("<br>")[0].Split(':')[1].Trim();

                    questionAnswer.CompetencyName = "MCQ Questions And Answers";
                    questionAnswer.TagName = headerArray[k];
                    //header.InnerText.Replace("MCQ Questions And Answers", " ");

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
                        singleMultipleQuestionsOptions.OptionDetail = item.Split('.')[1].Trim();
                        singleMultipleQuestionsOptions.IsTrue = item.Split('.')[0].Trim() == questionAnswerList[i].Answer ? "TRUE" : "FALSE";

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
                    Tuple<string, MemoryStream> fileData = exportToExcelRepository.CreateExcelFileWithMultipleTable(dynamicDictionary, headerArray[k]);

                }
                catch (Exception)
                {
                    throw;
                }
                #endregion
            }
        }



    }
}
