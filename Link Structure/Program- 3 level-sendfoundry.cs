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
            HtmlDocument mainDocument = web.Load("https://www.sanfoundry.com/mechanical-engineering-questions-answers/");

            HtmlNode[] mainLinks = mainDocument.DocumentNode.SelectNodes("//div[@class='inside-article']//div[@class='entry-content']//table//tr//td//li//a").ToArray();

            //main links
            List<string> mianLinkArray = new List<string>();
            List<string> folderNames = new List<string>();
            //foreach (HtmlNode item in minLinks)
            for (int i = 0; i < mainLinks.Count(); i = i + 2)
            {
                HtmlAttribute att = mainLinks[i].Attributes["href"];
                if (att != null)
                {
                    Console.WriteLine(mainLinks[i].InnerText);
                    //Console.WriteLine(att.Value);
                    mianLinkArray.Add(att.Value);
                    folderNames.Add(mainLinks[i].InnerText);
                }
                //linkArray.Add("");

            }


            #region 3 level

            for (int a = 34; a < mianLinkArray.Count; a++)
            {
                string folderName = folderNames[a].Replace("&#038;", "&").ToString();
                string path = "c:\\khushbu\\Apptitude\\09-06-2021\\Mechanical" + "\\" + folderName;
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                #region 2 level
                //HtmlDocument document = web.Load("https://www.sanfoundry.com/1000-thermal-engineering-questions-answers/");

                HtmlDocument document = web.Load(mianLinkArray[a]);

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
                }

                string fileName = @"c:\\khushbu\\Apptitude\\09-06-2021\\Mechanical" + "\\" + folderName + "\\" + folderName + ".txt";
                //string fileName = @"C:\Temp\MaheshTX.txt";

                //// Check if file already exists. If yes, delete it.     
                //if (File.Exists(fileName))
                //{
                //    File.Delete(fileName);
                //}

                // Create a new file     
                using (StreamWriter sw = File.CreateText(fileName))
                {
                    foreach (var item in linkArray)
                    {

                        //sw.WriteLine("Add one more line ");
                        sw.WriteLine(item);
                    }
                }

                string tagName = "";
                for (int k = 0; k < linkArray.Count; k++)
                {
                    List<QuestionAnswer> questionAnswerList = new List<QuestionAnswer>();

                    HtmlDocument quesionDocument = web.Load(linkArray[k]);
                    HtmlNode entry = quesionDocument.DocumentNode.SelectNodes("//div[@class='entry-content']").First();

                    HtmlNode header = quesionDocument.DocumentNode.SelectNodes("//h1[@class='entry-title']").First();

                    #region set question and options
                    HtmlNode[] questions = entry.SelectNodes("//p").Where(x => !x.InnerHtml.Contains("strong")).Skip(1).SkipLast(1).ToArray();
                    HtmlNode[] answers = entry.SelectNodes("//div[@class='collapseomatic_content ']").ToArray();

                    string headerText = header.InnerText.Replace("&#8217;", "'").Replace("/", " ");
                    string[] headerData = headerText.Split(" &#8211;");

                    for (int i = 0; i < questions.Count(); i++)
                    {
                        QuestionAnswer questionAnswer = new QuestionAnswer();

                        questionAnswer.Options = new List<string>();
                        if (questions[i].SelectNodes("span") == null)
                        {
                            var questionText = questions[i].InnerText.Split("\n")[0];
                            for (int l = 1; l < questionText.Split('.').Length; l++)
                            {
                                questionText.Replace("&#8217;", "'").ToString();
                                questionText.Replace("&#8211;", "-").ToString();
                                questionText.Replace("&#038;", "&").ToString();
                                questionText.Replace("&lt;", "<").ToString();
                                questionText.Replace("&gt;", ">").ToString();
                                questionText.Replace("&#8220;", "'").ToString();
                                questionText.Replace("&#8221;", "'").ToString();

                                questionAnswer.QuestionDetails = questionAnswer.QuestionDetails + questionText.Split(".")[l];
                            }

                            i = i + 1;
                            if (i < questions.Count())
                            {
                                var options = questions[i].InnerText.Split("\n").SkipLast(1).ToArray();
                                for (int m = 0; m < options.Count(); m++)
                                {
                                    options[m] = options[m].Replace("&#8211;", "-").ToString();
                                    options[m] = options[m].Replace("&#038;", "&").ToString();
                                    options[m] = options[m].Replace("&#8217;", "'").ToString();
                                    options[m] = options[m].Replace("&lt;", "<").ToString();
                                    options[m] = options[m].Replace("&gt;", ">").ToString();
                                    options[m] = options[m].Replace("&#8220;", "'").ToString();
                                    options[m] = options[m].Replace("&#8221;", "'").ToString();

                                    questionAnswer.Options.Add(options[m]);
                                }
                            }
                        }

                        else
                        {
                            var questionText = questions[i].InnerText.Split("\n")[0];
                            for (int l = 1; l < questionText.Split('.').Length; l++)
                            {
                                questionText.Replace("&#8217;", "'").ToString();
                                questionText.Replace("&#8211;", "-").ToString();
                                questionText.Replace("&#038;", "&").ToString();
                                questionText.Replace("&lt;", "<").ToString();
                                questionText.Replace("&gt;", ">").ToString();
                                questionText.Replace("&#8220;", "'").ToString();
                                questionText.Replace("&#8221;", "'").ToString();
                                questionAnswer.QuestionDetails = questionAnswer.QuestionDetails + questionText.Split(".")[l];
                            }

                            var options = questions[i].InnerText.Split("\n").Skip(1).SkipLast(1).ToArray();
                            for (int m = 0; m < options.Count(); m++)
                            {
                                options[m] = options[m].Replace("&#8211;", "-").ToString();
                                options[m] = options[m].Replace("&#038;", "&").ToString();
                                options[m] = options[m].Replace("&lt;", "<").ToString();
                                options[m] = options[m].Replace("&gt;", ">").ToString();
                                options[m] = options[m].Replace("&#8217;", "'").ToString();
                                options[m] = options[m].Replace("&#8220;", "'").ToString();
                                options[m] = options[m].Replace("&#8221;", "'").ToString();

                                questionAnswer.Options.Add(options[m]);
                            }
                        }

                        questionAnswer.CompetencyName = headerData[0].Replace("&#038;", "&").Trim();
                        tagName = questionAnswer.TagName = headerData.Length > 1 ? headerData[1].Replace("&#038;", "&") : headerData[0].Replace("&#038;", "&");
                        if (questionAnswer.QuestionDetails != string.Empty)
                        {
                            questionAnswerList.Add(questionAnswer);

                        }

                    }


                    //answer
                    for (int i = 0; i < questionAnswerList.Count(); i++)
                    {
                        questionAnswerList[i].QuestionId = (i + 1).ToString();
                        questionAnswerList[i].QuestionType = "Single";
                        questionAnswerList[i].DifficultyLevel = "Medium";

                        questionAnswerList[i].TagName = questionAnswerList[i].TagName;
                        questionAnswerList[i].CompetencyName = questionAnswerList[i].CompetencyName;
                        questionAnswerList[i].Answer = answers[i].InnerText.Split("\n")[0].Replace("Answer:", "").Trim();
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
                            singleMultipleQuestionsOptions.OptionDetail = item.Replace("a)", "").Replace("b)", "").Replace("c)", "").Replace("d)", "");
                            singleMultipleQuestionsOptions.IsTrue = item.Split(")")[0].ToString() == questionAnswerList[i].Answer ? "TRUE" : "FALSE";

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
                        int id = k + 1;

                        Tuple<string, MemoryStream> fileData = exportToExcelRepository.CreateExcelFileWithMultipleTable(dynamicDictionary, tagName + "-" + id, path);

                    }
                    catch (Exception e)
                    {
                        throw;
                    }
                    #endregion
                }

                #endregion

            }
            #endregion
        }
    }
}
