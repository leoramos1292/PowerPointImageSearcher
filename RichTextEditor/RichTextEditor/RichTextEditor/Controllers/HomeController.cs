using RichTextEditor.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;

namespace RichTextEditor.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        } 

        List<string> searchTerms = new List<string>();

        [HttpPost]
        public ActionResult SelectImage(RichTextEditorViewModel slide)
        {
            RichTextEditorViewModel model = new RichTextEditorViewModel();
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

            model.ImagePaths = new List<string>();

            DirectoryInfo d = new DirectoryInfo(@"C:\Users\lazlo\Desktop\GitHub\PowerPointImageSearcher\RichTextEditor\RichTextEditor\RichTextEditor\images");
            FileInfo[] Files = d.GetFiles("*.png"); //Getting png files
            string str = "";
            foreach (FileInfo file in Files)
            {
                str = file.Name;
                model.ImagePaths.Add(str);
            }


            return View(model);
        }

        [HttpPost]
        public ActionResult GeneratePowerpoint(string[] imageCheckBox)
        {
            RichTextEditorViewModel model = new RichTextEditorViewModel();
            model.ImagePaths = new List<string>();

            for (int i = 0; i < imageCheckBox.Length; i++)
            {
                model.ImagePaths.Add(imageCheckBox[i]);
            }
            
            return View(model);
        }
        
    }
}