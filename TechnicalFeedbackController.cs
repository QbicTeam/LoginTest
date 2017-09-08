using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Cors;
using DHL.Chango.Core;
using DHL.Chango.DataTypes.DTOs.Feedbacks;
using System.IO;
using System.Xml;
using System.Text;
using DHL.Chango.DataTypes.Entities;

// For more information on enabling Web API for empty projects, visit http://go.microsoft.com/fwlink/?LinkID=397860

namespace DHL.ChangoInfo.API.Controllers
{
    [EnableCors("SiteCorsPolicy")]
    [Route("api/technicalfeedbacks")]
    public class TechnicalFeedbackController : Controller
    {
        private TechnicalFeedbackManager _tfManager;

        private const string FEEDBACK_DATE = "[FeedbackDate]";
        private const string FEEDBACK_INTERVIEWER_NAME = "[InterviewerName]";
        private const string FEEDBACK_PRACTICE = "[practice]";
        private const string FEEDBACK_PRACTICE_ROLE = "[PracticeRole]";
        private const string FEEDBACK_SENORITY = "[Seniority]";
        private const string FEEDBACK_ADMINISTRATIVE_POSITION = "[AdministrativePosition]";
        private const string FEEDBACK_PROJECT_LEADER_INFO = "[ProjectLeaderInfo]";

        private const string FEEDBACK_THEORICAL_RESULT = "[TheoricalResult]";
        private const string FEEDBACK_THEORICAL_COMMENTS = "[TheoricalComments]";
        private const string FEEDBACK_PRACTICAL_RESULT = "[PracticalResult]";
        private const string FEEDBACK_PRACTICAL_COMMENTS = "[PracticalComments]";
        private const string FEEDBACK_POTENTIAL_RESULT = "[PotentialResult]";
        private const string FEEDBACK_POTENTIAL_COMMENTS = "[PotentialComments]";
        private const string FEEDBACK_COMMUNICATION_RESULT = "[CommunicationResult]";
        private const string FEEDBACK_COMMUNICATION_COMMENTS = "[CommunicationComments]";
        private const string FEEDBACK_INTEREST_RESULT = "[InterestResult]";
        private const string FEEDBACK_INTEREST_COMMENTS = "[InterestComments]";

        private const string FEEDBACK_ENGLISH_INTERVIEWER_NAME = "[EnglishInterviewerName]";
        private const string FEEDBACK_ENGLISH_VALIDATION_RESULT = "[EnglishValidationResult]";
        private const string FEEDBACK_FINAL_RESULT = "[FeedbackFinalResult]";
        private const string FEEDBACK_ADDITIONAL_COMMENTS = "[FinalAdditionalComments]";

        private const string FEEDBACK_TECHNICAL_CONTENT = "[TECHNICAL_CONTENT]";
        private const string FEEDBACK_TECHNICAL_ROW_CATEGORIE = "[FEEDBACK_TECH_SKILL_HEADER]";
        private const string FEEDBACK_TECHNICAL_ROW_SKILL = "[FEEDBACK_TECH_SKILL]";
        private const string FEEDBACK_TECHNICAL_ROW_EXPERTISE = "[FEEDBACK_TECH_EXPERTISE]";
        private const string FEEDBACK_TECHNICAL_ROW_COMMENTS = "[FEEDBACK_TECH_COMMENTS]";


        private const string FEEDBACK_FOREIGN_LANGUAGE_INTERVIEWER_ID = "";


        public TechnicalFeedbackController(TechnicalFeedbackManager tfManager)
        {
            this._tfManager = tfManager;
        }

        [HttpGet("feedback/{id}")]
        public IActionResult Get(int id)
        {
            var rst = this._tfManager.GetTechnicalFeedbackById(id);

            var result = new TechnicalFeedbackDTO();

            result.ID = rst.ID;
            result.FullProjectName = "(" + rst.Project.ShortName + ") " + rst.Project.FullName;
            result.FeedbackDate = rst.FeedbackDate.ToString("MMM.dd.yyyy");
            result.RoleName = rst.Role.Name;
            result.InterviewerName = rst.Interviewer.DisplayName;
            result.FinalRecomendationResult = rst.VeredictResult.ToString();

            result.EnglishValidationResult = rst.EnglishVeredictResult.ToString();
            result.AdditionalComments = rst.AdditionalComments;
            result.SkillCategories = new List<FeedbackSkillCategoryDTO>();

            //Fill Skill categories
            foreach (var cat in rst.SkillFeedbacks)
            {
                result.SkillCategories.Add(new FeedbackSkillCategoryDTO
                {
                    Name = cat.SkillName,
                    Skills = new List<FeedbackSkillDTO>()
                });

                foreach (var sk in cat.SkillFeedbacks)
                {
                    result.SkillCategories[result.SkillCategories.Count - 1].Skills.Add(new FeedbackSkillDTO
                    {
                        SkillName = sk.SkillName,
                        Result = sk.LevelObtained.ToString(),
                        Comments = sk.Comments
                    });
                }
            }

            //Fill concepts
            result.Concepts = new List<FeedbackConceptDTO>();
            foreach (var c in rst.ConceptsFeedbacks)
            {
                result.Concepts.Add(new FeedbackConceptDTO
                {
                    Concept = c.Concept.ToString(),
                    Result = c.LevelObtained.ToString(),
                    Comments = c.Comments
                });
            }

            return Ok(result);
        }

        //[Produces("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")]
        [Produces("application/vnd.ms-excel")]
        [HttpGet("feedback/{id}/{filename}")]
        public string DownloadTechnicalFeedback(int id, string filename)
        {
            var rst = this._tfManager.GetTechnicalFeedbackById(id);

            var result = CreateExcellFeedback(rst);

            return result;
            //return Ok(result);
        }


        private string CreateExcellFeedback(TechnicalFeedback feedback)
        {
            string result = string.Empty;
            string template = Directory.GetCurrentDirectory() + "\\templates\\FeedTechTemplate.xls";

            result = LoadXMLTemplate(template);
            result = result.Replace(FEEDBACK_DATE, feedback.FeedbackDate.ToString("MMM dd, yyyy"));
            result = result.Replace(FEEDBACK_INTERVIEWER_NAME, feedback.Interviewer.DisplayName);
            result = result.Replace(FEEDBACK_PRACTICE, feedback.Practice.Name);
            result = result.Replace(FEEDBACK_PRACTICE_ROLE, feedback.Practice.Roles[0].Description);
            result = result.Replace(FEEDBACK_SENORITY, feedback.ProfessionalLevel.ToString());
            result = result.Replace(FEEDBACK_ADMINISTRATIVE_POSITION, feedback.AdministrativePosition);
            result = result.Replace(FEEDBACK_PROJECT_LEADER_INFO, feedback.ProjectLeaderInfo);

            result = result.Replace(FEEDBACK_THEORICAL_RESULT, feedback.ConceptsFeedbacks[0].LevelObtained.ToString());
            result = result.Replace(FEEDBACK_THEORICAL_COMMENTS, feedback.ConceptsFeedbacks[0].Comments);
            result = result.Replace(FEEDBACK_PRACTICAL_RESULT, feedback.ConceptsFeedbacks[1].LevelObtained.ToString());
            result = result.Replace(FEEDBACK_PRACTICAL_COMMENTS, feedback.ConceptsFeedbacks[1].Comments);
            result = result.Replace(FEEDBACK_POTENTIAL_RESULT, feedback.ConceptsFeedbacks[2].LevelObtained.ToString());
            result = result.Replace(FEEDBACK_POTENTIAL_COMMENTS, feedback.ConceptsFeedbacks[2].Comments);
            result = result.Replace(FEEDBACK_COMMUNICATION_RESULT, feedback.ConceptsFeedbacks[3].LevelObtained.ToString());
            result = result.Replace(FEEDBACK_COMMUNICATION_COMMENTS, feedback.ConceptsFeedbacks[3].Comments);
            result = result.Replace(FEEDBACK_INTEREST_RESULT, feedback.ConceptsFeedbacks[4].LevelObtained.ToString());
            result = result.Replace(FEEDBACK_INTEREST_COMMENTS, feedback.ConceptsFeedbacks[4].Comments);

            result = result.Replace(FEEDBACK_ENGLISH_INTERVIEWER_NAME, feedback.ForeignLanguageInterviewer.DisplayName); 
            result = result.Replace(FEEDBACK_ENGLISH_VALIDATION_RESULT, feedback.EnglishVeredictResult.ToString());
            result = result.Replace(FEEDBACK_INTERVIEWER_NAME, feedback.Interviewer.DisplayName);
            result = result.Replace(FEEDBACK_FINAL_RESULT, feedback.VeredictResult.ToString());
            result = result.Replace(FEEDBACK_ADDITIONAL_COMMENTS, feedback.AdditionalComments);

            string techContent = string.Empty;

            foreach(var itm in feedback.SkillFeedbacks)
            {
                techContent += GetTechnicalHeader(itm.SkillName);

                //Details
                foreach (var dt in itm.SkillFeedbacks)
                {
                    techContent += GetThechicalDataRow(dt.SkillName, dt.LevelObtained.ToString(), dt.Comments);
                }

                //Footer...
                techContent += GetBlankRow();
            }

            result = result.Replace(FEEDBACK_TECHNICAL_CONTENT, techContent);


            return result;
        }

        private string GetThechicalDataRow(string technology, string level, string comments)
        {
            string result = string.Empty;
            result = "<Row ss:AutoFitHeight=\"0\" ss:Height=\"15\" ss:StyleID=\"s165\"><Cell ss:Index=\"2\" ss:MergeAcross=\"4\" ss:StyleID=\"m85263904\"><Data ss:Type=\"String\">[FEEDBACK_TECH_SKILL]</Data></Cell><Cell ss:MergeAcross=\"6\" ss:StyleID=\"m85263924\"><Data ss:Type=\"String\">[FEEDBACK_TECH_EXPERTISE]</Data></Cell><Cell ss:MergeAcross=\"16\" ss:StyleID=\"m85263944\"><Data ss:Type=\"String\">[FEEDBACK_TECH_COMMENTS]</Data></Cell><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/></Row>";

            result = result.Replace(FEEDBACK_TECHNICAL_ROW_SKILL, technology);
            result = result.Replace(FEEDBACK_TECHNICAL_ROW_EXPERTISE, level);
            result = result.Replace(FEEDBACK_TECHNICAL_ROW_COMMENTS, comments);

            return result;
        }

        private string GetTechnicalHeader(string header)
        {
            string result = string.Empty;
            result = "<Row ss:AutoFitHeight=\"0\" ss:Height=\"12.9375\" ss:StyleID=\"s63\"><Cell ss:Index=\"2\" ss:MergeAcross=\"4\" ss:StyleID=\"m85263516\"><Data ss:Type=\"String\">[FEEDBACK_TECH_SKILL_HEADER]</Data></Cell><Cell ss:MergeAcross=\"6\" ss:StyleID=\"m85263536\"><Data ss:Type=\"String\">LEVEL OBTAINED IN EVALUATION</Data></Cell><Cell ss:MergeAcross=\"16\" ss:StyleID=\"m85263556\"><Data ss:Type=\"String\">COMMENTS (Indicate certifications)</Data></Cell></Row>";

            result = result.Replace(FEEDBACK_TECHNICAL_ROW_CATEGORIE, header);

            return result;
        }

        private string GetBlankRow()
        {
            string result = string.Empty;

            result = "<Row ss:AutoFitHeight=\"0\" ss:Height=\"12\" ss:StyleID=\"s165\"><Cell ss:Index=\"2\" ss:StyleID=\"s186\"/><Cell ss:StyleID=\"s148\"/><Cell ss:StyleID=\"s148\"/><Cell ss:StyleID=\"s148\"/><Cell ss:StyleID=\"s148\"/><Cell ss:StyleID=\"s187\"/><Cell ss:StyleID=\"s148\"/><Cell ss:StyleID=\"s148\"/><Cell ss:StyleID=\"s148\"/><Cell ss:StyleID=\"s148\"/><Cell ss:StyleID=\"s148\"/><Cell ss:StyleID=\"s148\"/><Cell ss:StyleID=\"s188\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s189\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/><Cell ss:StyleID=\"s185\"/></Row>";

            return result;
        }

        private string LoadXMLTemplate(string path)
        {
            string result = string.Empty;

            FileStream fs = new FileStream(path, FileMode.Open);
            using (StreamReader reader = new StreamReader(fs))
            {
                result = reader.ReadToEnd();
            }

            return result;
        }


    }
}

