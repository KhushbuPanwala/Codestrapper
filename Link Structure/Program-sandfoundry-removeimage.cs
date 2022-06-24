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
            HtmlDocument mainDocument = web.Load("https://www.sanfoundry.com/master-computer-applications-questions-answers/");

            HtmlNode[] mainLinks = mainDocument.DocumentNode.SelectNodes("//div[@class='inside-article']//div[@class='entry-content']//table//tr//td//li//a").ToArray();

            //main links
            List<string> mainLinkArray = new List<string>();
            List<string> folderNames = new List<string>();
            for (int i = 0; i < mainLinks.Count(); i++)
            {
                HtmlAttribute att = mainLinks[i].Attributes["href"];
                if (att != null)
                {
                    //Console.WriteLine(mainLinks[i].InnerText);
                    //Console.WriteLine(att.Value);
                    if (!att.Value.Contains("test"))
                    {
                        mainLinkArray.Add(att.Value);
                        folderNames.Add(mainLinks[i].InnerText);
                    }
                }
            }


            #region 3 level

            for (int a = 0; a < mainLinkArray.Count(); a++)
            //for (int a = 0; a < 13; a++)
            {
                string folderName = folderNames[a].Replace("&#038;", "&").Replace("&#8211;", "-").Replace("/", " ").ToString();
                string path = "c:\\khushbu\\Apptitude\\09-06-2021\\master-computer-applications" + "\\" + folderName;
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                #region 2 level
                //HtmlDocument document = web.Load("https://www.sanfoundry.com/1000-thermal-engineering-questions-answers/");
                HtmlDocument document = web.Load(mainLinkArray[a]);

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

                //string fileName = @"c:\\khushbu\\Apptitude\\09-06-2021\\Mechanical" + "\\" + folderName + "\\" + folderName + ".txt";
                string fileName = path + "\\" + folderName + ".txt";

                // Create a new file     
                using (StreamWriter sw = File.CreateText(fileName))
                {
                    foreach (var item in linkArray)
                    {
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
                        if (!questions[i].InnerHtml.Contains("img"))
                        {
                            QuestionAnswer questionAnswer = new QuestionAnswer();

                            questionAnswer.Options = new List<string>();
                            if (questions[i].SelectNodes("span") == null)
                            {
                            }

                            else
                            {
                                string questionText = "";
                                for (int ik = 0; ik < questions[i].InnerText.Split("\n").Length - 1; ik++)
                                {
                                    if (questions[i].InnerText.Split("\n")[ik].Contains("a)") || questions[i].InnerText.Split("\n")[ik].Contains("b)") ||
                                        questions[i].InnerText.Split("\n")[ik].Contains("c)") || questions[i].InnerText.Split("\n")[ik].Contains("d)"))
                                    {
                                        // add in option
                                        string option = questions[i].InnerText.Split("\n")[ik];

                                        option = option.Replace("&#8211;", "-").Replace("&#038;", "&").ToString()
                                            .Replace("&#8217;", "'").Replace("&gt;", ">").Replace("&lt;", "<").ToString();

                                        questionAnswer.Options.Add(option);
                                    }
                                    else
                                    {
                                        questionText = questionText + " " + questions[i].InnerText.Split("\n")[ik];
                                        questionText = questionText.Replace("&#8211;", "-").Replace("&#8217;", "'")
                                        .Replace("&#8211;", "-").Replace("&#038;", "&").Replace("&gt;", ">")
                                        .Replace("&lt;", "<").ToString();

                                    }
                                }
                                for (int l = 1; l < questionText.Split('.').Length; l++)
                                {
                                    questionAnswer.QuestionDetails = questionAnswer.QuestionDetails + questionText.Split(".")[l];
                                }

                            }

                            questionAnswer.CompetencyName = headerData[0].Replace("&#038;", "&").Trim();
                            tagName = questionAnswer.TagName = headerData.Length > 1 ? headerData[1] : headerData[0];

                            questionAnswer.QuestionType = "Single";
                            questionAnswer.DifficultyLevel = "Medium";

                            questionAnswer.Answer = answers[i].InnerText.Split("\n")[0].Replace("Answer:", "").Trim();

                            if (!string.IsNullOrEmpty(questionAnswer.QuestionDetails))
                            {
                                questionAnswerList.Add(questionAnswer);

                            }
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
                    List<Instructions> instructionList = new List<Instructions>();
                    Instructions instructions = new Instructions();

                    instructions.Title = "Instruction Sheet";

                    instructionList.Add(instructions);

                    //create dynamic directory
                    dynamic dynamicDictionary = new DynamicDictionary<string, dynamic>();
                    dynamicDictionary.Add("Instruction", instructionList);
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
