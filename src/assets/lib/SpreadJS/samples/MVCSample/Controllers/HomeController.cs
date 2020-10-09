using MVCSample.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MVCSample.Controllers
{
    public class HomeController : Controller
    {
        private Entities db = new Entities();

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            return View();
        }

        public JsonResult Load(int index = 0, int count = 200)
        {
            this.HttpContext.Response.Expires = 0;

            int totalRecords = db.People.Count();

            IQueryable<Person> query = db.People.OrderBy(t => t.BusinessEntityID).Skip(index).Take(count);

            int n = 0;
            if (0 <= index && index < totalRecords)
            {
                n = Math.Min(count, totalRecords - index);
            }

            return Json(new
            {
                start = index,
                count = n,
                total = totalRecords,
                data = query,
            }, JsonRequestBehavior.AllowGet);
        }

        public JsonResult LoadAll(int index = 0)
        {
            this.HttpContext.Response.Expires = 0;

            int totalRecords = db.People.Count();

            IQueryable<Person> query = db.People.OrderBy(t => t.BusinessEntityID).Skip(index).Take(totalRecords);

            return Json(new
            {
                start = index,
                count = totalRecords,
                total = totalRecords,
                data = query,
            }, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult Create(Person[] contact)
        {
            try
            {
                int maxId = 0;
                IQueryable<Person> max = db.People.OrderByDescending(t => t.BusinessEntityID).Take(1);
                foreach (Person p in max)
                {
                    maxId = p.BusinessEntityID;
                }
                foreach (Person ps in contact)
                {
                    ps.ModifiedDate = DateTime.Now;
                    ps.BusinessEntityID = ++maxId;
                    db.People.Add(ps);
                }
                db.SaveChanges();
            }
            catch (Exception ex)
            {
                return Json(new
                {
                    data = "",
                    success = false,
                    message = ex.Message
                });
            }

            return Json(new
            {
                data = contact,
                success = true,
                message = "Create method called successfully"
            });
        }

        [HttpPost]
        public JsonResult Delete(Person contact)
        {
            Person t = db.People.Single(c => c.BusinessEntityID == contact.BusinessEntityID);
            db.People.Remove(t);
            db.SaveChanges();

            return Json(new
            {
                data = contact,
                success = true,
                message = "Delete method called successfully"
            });
        }

        [HttpPost]
        public JsonResult Update(Person contact)
        {
            try
            {
                contact.ModifiedDate = DateTime.Now;
                db.People.Attach(contact);
                db.Entry(contact).State = EntityState.Modified;
                db.SaveChanges();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return Json(new
            {
                data = contact,
                success = true,
                message = "Update method called successfully"
            });
        }

        [HttpPost]
        public JsonResult UpdateAll(Person[] contact)
        {
            try
            {
                foreach (Person ps in contact)
                {
                    ps.ModifiedDate = DateTime.Now;
                    db.People.Attach(ps);
                    db.Entry(ps).State = EntityState.Modified;
                }
                db.SaveChanges();
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return Json(new
            {
                data = contact,
                success = true,
                message = "Update method called successfully"
            });
        }

        public JsonResult Sort()
        {
            this.HttpContext.Response.Expires = 0;

            int totalRecords = db.People.Count();

            IQueryable<Person> query = db.People.OrderByDescending(t => t.FirstName).Take(totalRecords);

            return Json(new
            {
                data = query,
            }, JsonRequestBehavior.AllowGet);
        }
    }
}