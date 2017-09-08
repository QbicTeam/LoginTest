using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Cors;
using DHL.Chango.Core;
using AutoMapper;
using DHL.Chango.DataTypes.DTOs.Technologies;

// For more information on enabling Web API for empty projects, visit http://go.microsoft.com/fwlink/?LinkID=397860

namespace DHL.ChangoInfo.API.Controllers
{
    [EnableCors("SiteCorsPolicy")]
    [Route("api/technologies")]
    public class TechnologiesController : Controller
    {
        private TechnologiesManager _manager;

        public TechnologiesController(TechnologiesManager manager)
        {
            this._manager = manager;
        }


        [HttpGet("catalog")]
        public IActionResult GetTechnologiesCatalog()
        {
            var data = this._manager.GetAllParentTechnologies();
            var result = Mapper.Map<IEnumerable<TechnologyCatalogDTO>>(data);

            return Ok(result);
        }

    }
}




