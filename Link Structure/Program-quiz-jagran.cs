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
            //List<string> linkArray = new List<string>();
            //foreach (HtmlNode item in links)
            //{
            //    HtmlAttribute att = item.Attributes["href"];
            //    linkArray.Add(att.Value);
            //}

            List<QuestionAnswer> questionAnswerList = new List<QuestionAnswer>();
            //for (int k = 1; k <= 2; k++)
            //{

            //string link = "https://www.sawaal.com/aptitude-reasoning/quantitative-aptitude-arithmetic-ability/alligation-or-mixture-questions-and-answers.html";
            //string link = "https://www.sawaal.com/aptitude-reasoning/quantitative-aptitude-arithmetic-ability/alligation-or-mixture-questions-and-answers.htm?page=2&sort=";
            //string link = "https://www.avatto.com/ugc-net-paper1/paper1/mcqs/teaching-aptitude/questions/515/1.html";
            //https://quiz.jagranjosh.com/josh/quiz/index.php?attempt_id=10576951&page=1
            //+ k;
            string link = "https://quiz.jagranjosh.com/josh/quiz/index.php?attempt_id=10576951";
            for (int k = 1; k < 7; k++)
            {


                HtmlDocument quesionDocument = web.Load(link);

                HtmlNode entry1 = quesionDocument.DocumentNode.SelectSingleNode("//div[@class='cus-container']");
                string header = quesionDocument.DocumentNode.SelectSingleNode("//div[@class='page-heading']").InnerText.Trim();

                //HtmlNode[] entry1 = quesionDocument.DocumentNode.SelectNodes("//div[@class='quessect']").ToArray();

                //Console.Write(entry1);

                string question = entry1.SelectSingleNode("//div[@class='nquizbox onlinetest']").SelectSingleNode("ul").SelectNodes("li")[0].SelectSingleNode("p").InnerText;

                int questionCount = entry1.SelectSingleNode("//div[@class='nquizbox onlinetest']").SelectSingleNode("ul").SelectNodes("li").Count();
                for (int i = 0; i < questionCount; i++)
                {
                    QuestionAnswer questionAnswer = new QuestionAnswer();
                    questionAnswer.QuestionId = entry1.SelectSingleNode("//div[@class='nquizbox onlinetest']").SelectSingleNode("ul").SelectNodes("li")[i].SelectSingleNode("p").SelectSingleNode("strong").InnerText.Replace('Q', ' ').Replace('.', ' ').Trim();
                    questionAnswer.QuestionDetails = entry1.SelectSingleNode("//div[@class='nquizbox onlinetest']").SelectSingleNode("ul").SelectNodes("li")[i].SelectSingleNode("p").InnerText.Trim();
                    questionAnswer.QuestionType = "Single";
                    questionAnswer.DifficultyLevel = "Medium";
                    //questionAnswer.CompetencyName = header.Split('-')[0];
                    questionAnswer.TagName = header;
                    List<string> options = new List<string>();

                    int optionCount = entry1.SelectSingleNode("//div[@class='nquizbox onlinetest']").SelectSingleNode("ul").SelectNodes("li")[i].SelectSingleNode("ul").SelectNodes("div").Count();
                    //[0].InnerText.Trim();
                    for (int j = 0; j < optionCount; j++)
                    {
                        options.Add(entry1.SelectSingleNode("//div[@class='nquizbox onlinetest']").SelectSingleNode("ul").SelectNodes("li")[i].SelectSingleNode("ul").SelectNodes("div")[j].InnerText.Trim());

                    }
                    questionAnswer.Options = options;
                    questionAnswerList.Add(questionAnswer);
                }
                link = "https://quiz.jagranjosh.com/josh/quiz/index.php?attempt_id=10576951&page=" + k;

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

                for (int j = 0; j < questionAnswerList[i].Options.Count(); j++)
                {

                    singleMultipleQuestionsOptions = new SingleMultipleQuestionsOptions();

                    singleMultipleQuestionsOptions.QuestionId = questionAnswerList[i].QuestionId;
                    singleMultipleQuestionsOptions.OptionDetail = questionAnswerList[i].Options[j];

                    singleMultipleQuestionsOptions.IsTrue = "FALSE";
                    //singleMultipleQuestionsOptions.IsTrue = questionAnswerList[i].Answer == questionAnswerList[i].OptionsIds[j] ? "TRUE" : "FALSE";

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
