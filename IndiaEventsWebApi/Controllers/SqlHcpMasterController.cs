//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Threading.Tasks;
//using Microsoft.AspNetCore.Http;
//using Microsoft.AspNetCore.Mvc;
//using Microsoft.EntityFrameworkCore;
//using IndiaEventsWebApi;
//using IndiaEventsWebApi.Models.MasterSheets;

//namespace IndiaEventsWebApi.Controllers
//{
//    [Route("api/[controller]")]
//    [ApiController]
//    public class SqlHcpMasterController : ControllerBase
//    {
//        private readonly DataContext _context;

//        public SqlHcpMasterController(DataContext context)
//        {
//            _context = context;
//        }

//        [HttpGet]
//        public async Task<ActionResult<IEnumerable<HcpMasterData>>> GetHcpMaster()
//        {
//            return await _context.HCPMaster.ToListAsync();
//        }

//    }
//}
