using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;

namespace HTMLAgility
{
    class Program
    {
        static void Main(string[] args)
        {

            HtmlWeb web = new HtmlWeb();
            //HtmlDocument document = web.Load("https://www.sanfoundry.com/java-questions-answers-freshers-experienced/");
            ////links
            //HtmlNode[] links = document.DocumentNode.SelectNodes("//div[@class='sf-section']//table//tr//td//li//a").ToArray();
            List<string> linkArray = new List<string>();
            foreach (HtmlNode item in links)
            {
                HtmlAttribute att = item.Attributes["href"];
                linkArray.Add(att.Value);
            }
            List<QuestionAnswer> questionAnswerList = new List<QuestionAnswer>();

            ////for (int k = 0; k < linkArray.Count; k++)
            //for (int k = 0; k < 5; k++)
            //{
                HtmlDocument quesionDocument = web.Load(linkArray[0]);
                HtmlNode entry = quesionDocument.DocumentNode.SelectNodes("//div[@class='entry-content']").First();

                HtmlNode header = quesionDocument.DocumentNode.SelectNodes("//h1[@class='entry-title']").First();

                #region set question and options
                HtmlNode[] questions = entry.SelectNodes("//p").Where(x => !x.InnerHtml.Contains("strong")).Skip(1).SkipLast(1).ToArray();
                HtmlNode[] answers = entry.SelectNodes("//div[@class='collapseomatic_content ']").ToArray();

                string[] headerData = header.InnerHtml.Split("&#8211;");

                for (int i = 0; i < questions.Count(); i++)
                {
                    QuestionAnswer questionAnswer = new QuestionAnswer();

                    questionAnswer.Options = new List<string>();
                    if (questions[i].SelectNodes("span") == null)
                    {
                        questionAnswer.QuestionDetails = questions[i].InnerText.Split("\n")[0];
                        i = i + 1;
                        if (i < questions.Count())
                        {
                            var options = questions[i].InnerText.Split("\n").SkipLast(1).ToArray();
                            foreach (string opt in options)
                            {
                                questionAnswer.Options.Add(opt);
                            }
                        }
                    }

                    else
                    {
                        var questionText = questions[i].InnerText.Split("\n")[0];
                        questionAnswer.QuestionDetails = questions[i].InnerText.Split("\n")[0];

                        var options = questions[i].InnerText.Split("\n").Skip(1).SkipLast(1).ToArray();
                        foreach (string opt in options)
                        {
                            questionAnswer.Options.Add(opt);
                        }
                    }

                    questionAnswer.CompetencyName = headerData[0].Replace("&#038;", "&").Trim();
                    questionAnswer.TagName = headerData[1];


                    questionAnswerList.Add(questionAnswer);
                }


                //answer
                for (int i = 0; i < questionAnswerList.Count(); i++)
                {
                    questionAnswerList[i].QuestionId = (i + 1).ToString();
                    questionAnswerList[i].QuestionType = "Single";
                    questionAnswerList[i].DifficultyLevel = "Medium";

                    foreach (var item in questionAnswerList[i].Options)
                    {
                        questionAnswerList[i].QuestionDetails += " " + item;
                    }
                    questionAnswerList[i].TagName = questionAnswerList[i].TagName;
                    questionAnswerList[i].CompetencyName = questionAnswerList[i].CompetencyName;
                    questionAnswerList[i].Answer = answers[i].InnerText.Split("\n")[0].Replace("Answer:", "").Trim();
                }

                #endregion
            //}

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
                    singleMultipleQuestionsOptions.IsTrue = singleMultipleQuestionsOptions.OptionDetail.Split(")")[0].ToString() == questionAnswerList[i].Answer ? "TRUE" : "FALSE";

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
