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
            for (int k = 1; k <= 2; k++)
            {

                string link = "https://www.indiabix.com/aptitude/height-and-distance/06900" + k;


                HtmlDocument quesionDocument = web.Load(link);

                HtmlNode[] entry1 = quesionDocument.DocumentNode.SelectNodes("//div[@class='bix-div-container']").ToArray();


                string header = quesionDocument.DocumentNode.SelectNodes("//div[@class='pagehead']")[0].InnerText;


                for (int i = 0; i < entry1.Count(); i++)
                {
                    QuestionAnswer questionAnswer = new QuestionAnswer();
                    questionAnswer.QuestionId = entry1[i].SelectNodes("table")[0].SelectNodes("tr")[0].SelectNodes("td")[0].InnerText.Split(".")[0];

                    questionAnswer.QuestionDetails = entry1[i].SelectNodes("table")[0].SelectNodes("tr")[0].SelectNodes("td")[1].InnerText;

                    //options
                    List<string> options = new List<string>();
                    List<string> optionIds = new List<string>();

                    var optionsDataCount = entry1[i].SelectNodes("table")[0].SelectNodes("tr")[1].SelectNodes("td")[0].SelectNodes("table")[0].SelectNodes("tr").Count();
                    //var optionsData = entry1[i].SelectNodes("table")[0].SelectNodes("tr")[1].SelectNodes("td")[0].SelectNodes("table")[0].SelectNodes("tr")[i].InnerText.Split("\n")[1].Trim()
                    for (int j = 0; j < optionsDataCount; j++)
                    {
                        optionIds.Add(entry1[i].SelectNodes("table")[0].SelectNodes("tr")[1].SelectNodes("td")[0].SelectNodes("table")[0].SelectNodes("tr")[j].InnerText.Split(".")[0].Trim());
                        options.Add(entry1[i].SelectNodes("table")[0].SelectNodes("tr")[1].SelectNodes("td")[0].SelectNodes("table")[0].SelectNodes("tr")[j].InnerText.Split("\n")[1].Trim());
                    }

                    questionAnswer.OptionsIds = optionIds;
                    questionAnswer.Options = options;
                    questionAnswer.Answer = entry1[i].SelectNodes("table")[0].SelectNodes("tr")[1].SelectNodes("td")[0].SelectSingleNode("input").Attributes["value"].Value;
                    //questionAnswer.Answer = entry1[i].SelectNodes("table")[0].SelectNodes("tr")[1].SelectNodes("td")[0].SelectSingleNode("//input[@type='hidden' and @class='jq-hdnakq']").Attributes["value"].Value;
                    questionAnswer.CompetencyName = header.Split('-')[0];
                    questionAnswer.TagName = header.Split('-')[1];
                    questionAnswer.QuestionType = "Single";
                    questionAnswer.DifficultyLevel = "Medium";

                    questionAnswerList.Add(questionAnswer);
                }

            }
            //#endregion


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

                    singleMultipleQuestionsOptions.IsTrue = questionAnswerList[i].Answer == questionAnswerList[i].OptionsIds[j] ? "TRUE" : "FALSE";

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
