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
            HtmlDocument document = web.Load("https://www.sanfoundry.com/1000-html-questions-answers/");
            //links
            HtmlNode[] links = document.DocumentNode.SelectNodes("//div[@class='sf-section']//table//tr//td//li//a").ToArray();
            List<string> linkArray = new List<string>();
            foreach (HtmlNode item in links)
            {
                HtmlAttribute att = item.Attributes["href"];
                if (att != null)
                {
                    Console.WriteLine(att.Value);
                    linkArray.Add(att.Value);
                }
                //linkArray.Add("");

            }

            List<QuestionAnswer> questionAnswersDetail = new List<QuestionAnswer>();
            for (int k = 0; k < linkArray.Count; k++)
            {
                List<QuestionAnswer> questionAnswerList = new List<QuestionAnswer>();

                HtmlDocument quesionDocument = web.Load(linkArray[k]);
                //string link = "https://www.sanfoundry.com/data-structure-interview-questions-answers-experienced/";
                //HtmlDocument quesionDocument = web.Load(link);
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
                        //questionAnswer.QuestionId = questions[i].InnerText.Split('.')[0];
                        //questionAnswer.QuestionDetails = questions[i].InnerText.Split("\n")[0];
                        //for (int l = 1; l < questions[i].InnerText.Split('.').Length; l++)
                        //{
                        //    questionAnswer.QuestionDetails = questionAnswer.QuestionDetails + " ." + questions[i].InnerText.Split(".")[l];
                        //}

                        //questionAnswer.QuestionDetails = questions[i].InnerText.Split('\n')[0].Split(".")[1].Trim();
                        var questionText = questions[i].InnerText.Split("\n")[0];
                        for (int l = 1; l < questionText.Split('.').Length; l++)
                        {
                            questionText.Replace("&#8217;", "'").ToString();
                            questionText.Replace("&#8211;", "-").ToString();
                            questionText.Replace("&#038;", "&").ToString();
                            questionAnswer.QuestionDetails = questionAnswer.QuestionDetails + questionText.Split(".")[l];
                        }

                        //questionAnswer.QuestionDetails = questions[i].InnerText.Split("\n")[0];
                        i = i + 1;
                        if (i < questions.Count())
                        {
                            var options = questions[i].InnerText.Split("\n").SkipLast(1).ToArray();
                            //foreach (string opt in options)
                            //{
                            //    questionAnswer.Options.Add(opt);
                            //}
                            for (int m = 0; m < options.Count(); m++)
                            {
                                options[m] = options[m].Replace("&#8211;", "-").ToString();
                                options[m] = options[m].Replace("&#038;", "&").ToString();
                                options[m] = options[m].Replace("&#8217;", "'").ToString();
                                questionAnswer.Options.Add(options[m]);
                            }
                        }
                    }

                    else
                    {
                        var questionText = questions[i].InnerText.Split("\n")[0];
                        //questionAnswer.QuestionDetails = questions[i].InnerText.Split("\n")[0];
                        for (int l = 1; l < questionText.Split('.').Length; l++)
                        {
                            questionText.Replace("&#8217;", "'").ToString();
                            questionText.Replace("&#8211;", "-").ToString();
                            questionText.Replace("&#038;", "&").ToString();
                            questionAnswer.QuestionDetails = questionAnswer.QuestionDetails + questionText.Split(".")[l];
                        }

                        //questionAnswer.QuestionDetails = questions[i].InnerText.Split('\n')[0].Split(".")[1].Trim();
                        var options = questions[i].InnerText.Split("\n").Skip(1).SkipLast(1).ToArray();
                        //foreach (string opt in options)
                        for (int m = 0; m < options.Count(); m++)
                        {
                            options[m] = options[m].Replace("&#8211;", "-").ToString();
                            options[m] = options[m].Replace("&#038;", "&").ToString();
                            options[m] = options[m].Replace("&#8217;", "'").ToString();
                            questionAnswer.Options.Add(options[m]);
                        }
                    }

                    questionAnswer.CompetencyName = headerData[0].Replace("&#038;", "&").Trim();
                    questionAnswer.TagName = headerData.Length > 1 ? headerData[1] : headerData[0];

                    if (questionAnswer.QuestionDetails != string.Empty)
                    {
                        questionAnswerList.Add(questionAnswer);

                    }

                }


                //answer
                for (int i = 0; i < questionAnswerList.Count(); i++)
                {
                    //questionAnswerList[i].QuestionId = (i + 1).ToString();
                    questionAnswerList[i].QuestionType = "Single";
                    questionAnswerList[i].DifficultyLevel = "Medium";

                    //foreach (var item in questionAnswerList[i].Options)
                    //{
                    //    questionAnswerList[i].QuestionDetails += " " + item;
                    //}
                    questionAnswerList[i].TagName = questionAnswerList[i].TagName;
                    questionAnswerList[i].CompetencyName = questionAnswerList[i].CompetencyName;
                    questionAnswerList[i].Answer = answers[i].InnerText.Split("\n")[0].Replace("Answer:", "").Trim();
                }

                foreach (var item in questionAnswerList)
                {
                    questionAnswersDetail.Add(item);
                }
            }


            #endregion
            for (int i = 0; i < questionAnswersDetail.Count; i++)
            {
                questionAnswersDetail[i].QuestionId = (i + 1).ToString();
            }
            #region Bind data for excel
            Program generateExcel = new Program();
            List<Questions> questionDetails = new List<Questions>();
            for (int i = 0; i < questionAnswersDetail.Count; i++)
            {
                Questions questionSheet = new Questions();
                //questionSheet.QuestionId = (i + 1).ToString();
                questionSheet.QuestionId = questionAnswersDetail[i].QuestionId;
                questionSheet.QuestionType = questionAnswersDetail[i].QuestionType;
                questionSheet.DifficultyLevel = questionAnswersDetail[i].DifficultyLevel;
                questionSheet.QuestionDetails = questionAnswersDetail[i].QuestionDetails;
                questionSheet.BasicOrPremium = string.Empty;
                questionDetails.Add(questionSheet);
            }



            List<QuestionTags> questionTagDetails = new List<QuestionTags>();
            for (int i = 0; i < questionAnswersDetail.Count; i++)
            {
                QuestionTags questionTags = new QuestionTags();
                questionTags.QuestionId = questionAnswersDetail[i].QuestionId;
                questionTags.TagName = questionAnswersDetail[i].TagName;
                questionTags.CompetencyName = questionAnswersDetail[i].CompetencyName;
                questionTagDetails.Add(questionTags);
            }

            List<SingleMultipleQuestionsOptions> singleMultipleQuestionsOptionDetails = new List<SingleMultipleQuestionsOptions>();
            for (int i = 0; i < questionAnswersDetail.Count; i++)
            {
                SingleMultipleQuestionsOptions singleMultipleQuestionsOptions = new SingleMultipleQuestionsOptions();
                foreach (var item in questionAnswersDetail[i].Options)
                {
                    singleMultipleQuestionsOptions = new SingleMultipleQuestionsOptions();

                    singleMultipleQuestionsOptions.QuestionId = questionAnswersDetail[i].QuestionId;
                    //singleMultipleQuestionsOptions.OptionDetail = item.Split(")")[1].ToString();
                    singleMultipleQuestionsOptions.OptionDetail = item;
                    singleMultipleQuestionsOptions.IsTrue = item.Split(")")[0].ToString() == questionAnswersDetail[i].Answer ? "TRUE" : "FALSE";

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
            catch (Exception e)
            {
                throw;
            }
            #endregion
        }
    }
}
