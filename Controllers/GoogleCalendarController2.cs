using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace WebApplication1.Controllers
{
    public class GoogleCalendar2Controller : Controller
    {
        // GET: GoogleCalendarController
        public ActionResult Index()
        {
            return View();
        }

        // GET: GoogleCalendarController/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: GoogleCalendarController/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: GoogleCalendarController/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: GoogleCalendarController/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: GoogleCalendarController/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(int id, IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: GoogleCalendarController/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: GoogleCalendarController/Delete/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id, IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }
    }
}
