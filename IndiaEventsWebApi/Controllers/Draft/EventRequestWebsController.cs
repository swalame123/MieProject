//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Threading.Tasks;
//using Microsoft.AspNetCore.Http;
//using Microsoft.AspNetCore.Mvc;
//using Microsoft.EntityFrameworkCore;
//using IndiaEvents.Models.Models.SqlSampleCheckModel;
//using IndiaEventsWebApi;

//namespace IndiaEventsWebApi.Controllers.Draft
//{
//    [Route("api/[controller]")]
//    [ApiController]
//    public class EventRequestWebsController : ControllerBase
//    {
//        private readonly DataContext _context;

//        public EventRequestWebsController(DataContext context)
//        {
//            _context = context;
//        }

//        // GET: api/EventRequestWebs
//        //[HttpGet]
//        //public async Task<ActionResult<IEnumerable<EventRequestWeb>>> GetEventRequesrWeb()
//        //{
//        //    return await _context.EventRequesrWeb.ToListAsync();
//        //}

//        // GET: api/EventRequestWebs/5
//        [HttpGet("{id}")]
//        public async Task<ActionResult<EventRequestWeb>> GetEventRequestWeb(int id)
//        {
//            var eventRequestWeb = await _context.EventRequestsWeb.FindAsync(id);

//            if (eventRequestWeb == null)
//            {
//                return NotFound();
//            }

//            return eventRequestWeb;
//        }

//        // PUT: api/EventRequestWebs/5
//        // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
//        [HttpPut("{id}")]
//        public async Task<IActionResult> PutEventRequestWeb(int id, EventRequestWeb eventRequestWeb)
//        {
//            if (id != eventRequestWeb.Id)
//            {
//                return BadRequest();
//            }

//            _context.Entry(eventRequestWeb).State = EntityState.Modified;

//            try
//            {
//                await _context.SaveChangesAsync();
//            }
//            catch (DbUpdateConcurrencyException)
//            {
//                if (!EventRequestWebExists(id))
//                {
//                    return NotFound();
//                }
//                else
//                {
//                    throw;
//                }
//            }

//            return NoContent();
//        }

        
        
//        [HttpPost]
//        public async Task<ActionResult<EventRequestWeb>> PostEventRequestWeb(EventRequestWeb eventRequestWeb)
//        {
//            _context.EventRequestsWeb.Add(eventRequestWeb);
//            await _context.SaveChangesAsync();

//            return Ok("EventRequestWeb");
//        }

//        // DELETE: api/EventRequestWebs/5
//        [HttpDelete("{id}")]
//        public async Task<IActionResult> DeleteEventRequestWeb(int id)
//        {
//            var eventRequestWeb = await _context.EventRequestsWeb.FindAsync(id);
//            if (eventRequestWeb == null)
//            {
//                return NotFound();
//            }

//            _context.EventRequestsWeb.Remove(eventRequestWeb);
//            await _context.SaveChangesAsync();

//            return NoContent();
//        }

//        private bool EventRequestWebExists(int id)
//        {
//            return _context.EventRequestsWeb.Any(e => e.Id == id);
//        }
//    }
//}
