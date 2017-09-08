using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using DHL.Chango.Core;
using AutoMapper;
using DHL.Chango.DataTypes.DTOs.Dashboard;
using DHL.Chango.DataTypes.Enums;
using DHL.Chango.DataTypes.Entities;
using Microsoft.AspNetCore.Cors;
using DHL.Chango.DataTypes.DTOs.Projects;
using DHL.Chango.DataTypes.DTOs.Users;

// For more information on enabling Web API for empty projects, visit http://go.microsoft.com/fwlink/?LinkID=397860

namespace DHL.ChangoInfo.API.Controllers
{
    [EnableCors("SiteCorsPolicy")]
    [Route("api/users")]
    public class UsersController : Controller
    {
        private ProfilesManager _manager;
        private ProjectsManager _pManager;
        private UsersManager _uManager;
        private const int TOP_NOTIFICATIONS = -1;

        public UsersController(ProfilesManager manager, ProjectsManager pmanager, UsersManager umanager)
        {
            this._manager = manager;
            this._pManager = pmanager;
            this._uManager = umanager;
        }

        [HttpGet("catalog")]
        public IActionResult GetAdminUsersCatalog()
        {
            var catalog = this._uManager.GetUsersCatalog(ChangoUserType.Administrator);

            var result = Mapper.Map<IEnumerable<UserCatalogDTO>>(catalog);
            return Ok(result);
        }

        [HttpGet("{userid}/profiles")]
        public IActionResult GetProfilesByUserId(int userid)
        {

            var profiles = this._manager.GetProfilesByUserId(userid, ProfileStatusType.InProgress);
            List<DashboardProfileDTO> result = new List<DashboardProfileDTO>();

            foreach (var profile in profiles)
            {
                result.Add(new DashboardProfileDTO
                {
                    ID = profile.ID,
                    FullName = (profile.FirstName + " " + profile.LastName + " " + profile.SecondLastName + " ").Trim(),
                    Title = profile.Title,
                    PhotoUrl = profile.PhotoUrl,
                    ProjectShortName = profile.Project != null ? profile.Project.ShortName : "Undefined",
                    RoleName = profile.Project != null ? profile.Project.Roles.First().Name : "",
                    ExperienceLevel = profile.Project != null ? profile.Project.Roles.First().Positions.First().Position.ToString() : "",
                    CheckList = FillCheckListItemsDTO(profile.CheckList.Items),
                    Skills = FillSkillsDTO(profile.Technologies, profile.Project != null ? profile.Project.Roles.First().Positions: new List<WorkPosition>())
                });
            }

            return Ok(result);
        }

        [HttpGet("{userid}/projects")]
        public IActionResult GetProjectsByUserId(int userid)
        {

            var projects = this._pManager.GetAllProjectByUserId(userid);
            var result = Mapper.Map<IEnumerable<WorkProjectCatalogDTO>>(projects);

            return Ok(result);
        }



        private List<CheckListItemDTO> FillCheckListItemsDTO(List<ChecklistItem> items)
        {
            List<CheckListItemDTO> result = new List<CheckListItemDTO>();

            foreach (var itm in items)
            {
                result.Add(new CheckListItemDTO
                {
                    Icon = itm.Icon,
                    Description = itm.Description,
                    IsCompleted = itm.IsCompleted
                });
            }

            return result;
        }

        private List<SkillDTO> FillSkillsDTO(List<SkillTag> tags, List<WorkPosition> positions)
        {
            List<SkillDTO> result = new List<SkillDTO>();

            foreach (var t in tags)
            {
                if (!string.IsNullOrEmpty(t.DisplayName.Trim()))
                {
                    result.Add(new SkillDTO
                    {
                        Tag = t.DisplayName.Trim(),
                        YOE = t.YearsOfExperience,
                        DisplayClass = "label label-default"
                    });
                }
            }

            foreach (var p in positions)
            {
                var itm = result.Find(s => s.Tag.ToLower() == p.Skill.DisplayName.ToLower());

                if (itm != null)
                {
                    if (itm.YOE < p.RequiredExperienceTime) itm.DisplayClass = "label label-warning";
                    else itm.DisplayClass = "label label-success";
                }
                else
                {
                    if (!string.IsNullOrEmpty(p.Skill.DisplayName.Trim()))
                    {
                        result.Add(new SkillDTO
                        {
                            Tag = p.Skill.DisplayName,
                            DisplayClass = p.SkillIsRequired ? "label label-danger" : "label label-default"
                        });
                    }
                }
            }

            return result;
        }


    }
}
