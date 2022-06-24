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
            //HtmlDocument document = web.Load("https://scholarexpress.com/multiple-choice-questions-mcq-on-project-management/");
            ////links
            //HtmlNode[] links = document.DocumentNode.SelectNodes("//div[@class='entry-content clearfix']//p//strong").Skip(1).ToArray();


            List<string> linkArray = new List<string>();
            //return;
            //foreach (HtmlNode item in links)
            //{
            //    HtmlAttribute att = item.Attributes["href"];
            //    if (att != null)
            //    {
            //        Console.WriteLine(att.Value);
            //        linkArray.Add(att.Value);
            //    }
            linkArray.Add("https://scholarexpress.com/mcq-on-ms-office/");
            linkArray.Add("https://scholarexpress.com/mcq-on-ms-office/2/");
            linkArray.Add("https://scholarexpress.com/mcq-on-ms-office/3/");

            //}
            List<QuestionAnswer> questionAnswerDetail = new List<QuestionAnswer>();
            for (int k = 0; k < linkArray.Count; k++)
            {

                HtmlDocument quesionDocument = web.Load(linkArray[k]);
                //string link = "https://scholarexpress.com/mcq-on-basic-computer/";



                //HtmlDocument quesionDocument = web.Load(link);

                #region set question and options
                //links
                HtmlNode[] questions = quesionDocument.DocumentNode.SelectNodes("//div[@class='entry-content clearfix']//p//strong").Skip(1).SkipLast(2).ToArray();
                HtmlNode[] options = quesionDocument.DocumentNode.SelectNodes("//div[@class='entry-content clearfix']//p").Where(x => !x.InnerHtml.Contains("strong")).ToArray();
                List<string> questionOptions = new List<string>();
                foreach (var item in options)
                {
                    string opt = item.InnerText.Trim();
                    if (!string.IsNullOrEmpty(opt) && opt != "&nbsp;")
                    {
                        questionOptions.Add(opt);
                    }
                }

                HtmlNode answerText = quesionDocument.DocumentNode.SelectNodes("//div[@class='entry-content clearfix']//p//strong").Last();
                List<string> answers = new List<string>();
                for (int i = 0; i < answerText.InnerText.Split(',').Length; i++)
                {
                    answers.Add(answerText.InnerText.Split(',')[i]);
                }

                #endregion

                int start = 0;
                int end = 0;
                int j = 0;
                for (int i = 0; i < questions.Length; i++)
                {
                    string questionText = questions[i].InnerText.ToString().Trim();
                    QuestionAnswer questionAnswer = new QuestionAnswer();
                    questionAnswer.Options = new List<string>();

                    questionAnswer.QuestionId = questionText.Split('-')[0];
                    for (int l = 1; l < questionText.Split('-').Length; l++)
                    {
                        questionText.Replace("&#8217;", "'").ToString();
                        questionText.Replace("&#8211;", "-").ToString();
                        questionText.Replace("&#038;", "&").ToString();
                        questionText.Replace("&gt;", ">").ToString();
                        questionText.Replace("&lt;", "<").ToString();
                        questionAnswer.QuestionDetails = questionAnswer.QuestionDetails + questionText.Split("-")[l];
                    }


                    if (i == 0)
                    {

                        start = i;
                        end = start + 3;
                    }
                    else
                    {
                        start = end + 1;
                        end = start + 3;
                        j = start;
                    }
                    if (start < questionOptions.Count)
                    {
                        while (j <= end)
                        {
                            questionAnswer.Options.Add(questionOptions[j]);
                            j++;
                        }

                    }


                    questionAnswer.CompetencyName = "Multiple Choice Questions (MCQ)";
                    questionAnswer.TagName = "Basic Computer Awerness";

                    questionAnswer.QuestionType = "Single";
                    questionAnswer.DifficultyLevel = "Medium";


                    questionAnswer.Answer = answers[i].Split('-')[1];
                    //.InnerText.Split("\n")[0].Replace("Answer:", "").Trim();

                    questionAnswerDetail.Add(questionAnswer);

                }
            }

            #region Bind data for excel
            Program generateExcel = new Program();
            List<Questions> questionDetails = new List<Questions>();
            for (int i = 0; i < questionAnswerDetail.Count; i++)
            {
                Questions questionSheet = new Questions();
                //questionSheet.QuestionId = (i + 1).ToString();
                questionSheet.QuestionId = questionAnswerDetail[i].QuestionId;
                questionSheet.QuestionType = questionAnswerDetail[i].QuestionType;
                questionSheet.DifficultyLevel = questionAnswerDetail[i].DifficultyLevel;
                questionSheet.QuestionDetails = questionAnswerDetail[i].QuestionDetails;
                questionSheet.BasicOrPremium = string.Empty;
                questionDetails.Add(questionSheet);
            }

            List<QuestionTags> questionTagDetails = new List<QuestionTags>();
            for (int i = 0; i < questionAnswerDetail.Count; i++)
            {
                QuestionTags questionTags = new QuestionTags();
                questionTags.QuestionId = questionAnswerDetail[i].QuestionId;
                questionTags.TagName = questionAnswerDetail[i].TagName;
                questionTags.CompetencyName = questionAnswerDetail[i].CompetencyName;
                questionTagDetails.Add(questionTags);
            }

            List<SingleMultipleQuestionsOptions> singleMultipleQuestionsOptionDetails = new List<SingleMultipleQuestionsOptions>();
            for (int i = 0; i < questionAnswerDetail.Count; i++)
            {
                SingleMultipleQuestionsOptions singleMultipleQuestionsOptions = new SingleMultipleQuestionsOptions();
                foreach (var item in questionAnswerDetail[i].Options)
                {
                    singleMultipleQuestionsOptions = new SingleMultipleQuestionsOptions();

                    singleMultipleQuestionsOptions.QuestionId = questionAnswerDetail[i].QuestionId;
                    //singleMultipleQuestionsOptions.OptionDetail = item.Split(")")[1].ToString();
                    singleMultipleQuestionsOptions.OptionDetail = item.Split(')')[1].Trim();
                    string answer = questionAnswerDetail[i].Answer.Replace('(', ' ').Replace(')', ' ').Trim().ToString();

                    singleMultipleQuestionsOptions.IsTrue = item.Split(')')[0].Replace('(', ' ').Trim() == answer ? "TRUE" : "FALSE";

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

                Tuple<string, MemoryStream> fileData = exportToExcelRepository.CreateExcelFileWithMultipleTable(dynamicDictionary, "Basic-computer-awareness");

            }
            catch (Exception e)
            {
                throw;
            }
            #endregion


        }
    }
}
