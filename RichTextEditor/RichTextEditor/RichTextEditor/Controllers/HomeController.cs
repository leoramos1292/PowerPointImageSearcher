using RichTextEditor.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace RichTextEditor.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(RichTextEditorViewModel slide)
        {
            List<string> searchTerms = new List<string>();

            string titleText = slide.Title;
            if (titleText != null)
            {
                int tFrom = titleText.IndexOf("<p>") + "<p>".Length;
                int tTo = titleText.LastIndexOf("</p>");
                string tResult = titleText.Substring(tFrom, tTo - tFrom);
                searchTerms.Add(tResult);
            }

            string bodyText = slide.Message;
            //bool bold
            if (bodyText != null && bodyText.Contains("<b>"))
            {
                int bFrom = bodyText.IndexOf("<b>") + "<b>".Length;
                int bTo = bodyText.LastIndexOf("</b>");
                string bResult = bodyText.Substring(bFrom, bTo - bFrom);
                searchTerms.Add(bResult);
            }


            return View();
        }
    }
}