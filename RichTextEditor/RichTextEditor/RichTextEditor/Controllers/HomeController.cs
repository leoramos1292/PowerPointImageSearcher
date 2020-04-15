using RichTextEditor.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Mvc;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace RichTextEditor.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }


        //This method recieves input from the two text boxes in the index view, uses that input to do a google API call, and get google image links and puts them in a list
        [HttpPost]
        public async Task<ActionResult> SelectImage(RichTextEditorViewModel model)
        {
            
            try
            {
                model.ImagePaths = new List<string>();

                //this code block the the Title string as an html element and parses through it to remove html tags. It then adds it to a search terms list
                List<string> searchTerms = new List<string>();
                string titleText = model.Title;
                if (titleText != null)
                {
                    titleText = titleText.Replace("<p>", "");
                    titleText = titleText.Replace("</p>", "");
                    titleText = titleText.Replace("<b>", "");
                    titleText = titleText.Replace("</b>", "");
                    titleText = titleText.Replace(" ", "+");
                    searchTerms.Add(titleText);
                }

                //this code block gets the message string and parses through it to remove html tags. It also removes all text that isn't bold and adds those to the search terms list
                string bodyText = model.Message;
                while (bodyText != null && bodyText.Contains("<b>"))
                {
                    Tuple<string, int> foo = ExtractString(bodyText, "b");
                    searchTerms.Add(foo.Item1.Replace("&nbsp;", "").TrimStart(' ').TrimEnd(' ').Replace(" ", "+"));
                    bodyText = bodyText.Remove(0, foo.Item2);
                }

                string concat = String.Join("+", searchTerms.ToArray());

                //This code block take the search terms and does a google api call. It returns a Json object that then gets parsed through to find the link to the images. It then add those links to a string
                using (var client = new HttpClient())
                {

                    string url = "https://www.googleapis.com/customsearch/v1?key=AIzaSyB9c0EZ7nJeQvq8nzSM8rKykzO4tI56sp8&cx=010904202939473966172:gzjxnqlpfjo&q=&searchType=image&fileType=jpg&imgSize=small&alt=json";

                    client.DefaultRequestHeaders.Clear();
                    //Define request data format  
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    HttpResponseMessage Res = await client.GetAsync(url.Insert(126, concat));

                    //Checking the response is successful or not which is sent using HttpClient  
                    if (Res.IsSuccessStatusCode)
                    {
                        ////Storing the response details recieved from web api   
                        var result = Res.Content.ReadAsStringAsync().Result;
                        string jsonResult = Newtonsoft.Json.JsonConvert.DeserializeObject(result).ToString();
                        var details = JObject.Parse(jsonResult);
                        var items = details["items"];

                        foreach (var item in items)
                        {
                            model.ImagePaths.Add(item["link"].ToString().TrimStart('"').TrimEnd('"'));
                        }
                    }
                }
            }
            catch
            {
                ErrorPage();
            }
            
            return View(model);

        }

        //This method assist in parsing through the message text (2nd text box) that uses bold text as keywords.
        public Tuple<string, int> ExtractString(string s, string tag)
        {

            //try catch on this
            var startTag = "<" + tag + ">";
            int startIndex = s.IndexOf(startTag) + startTag.Length;
            int endIndex = s.IndexOf("</" + tag + ">", startIndex);

            Tuple<string, int> result = new Tuple<string, int>(s.Substring(startIndex, endIndex - startIndex), endIndex+4);
            return result;
        }
    

        //This Method receives the images with selected checkboxes and passes them to the GeneratePowerPointFile method which creates a powerpoint file using the images passed to it
        [HttpPost]
        public ActionResult GeneratePowerpoint(string[] imageCheckBox, RichTextEditorViewModel model)
        {
            model.ImagePaths = new List<string>();

            for (int i = 0; i < imageCheckBox.Length; i++)
            {
                model.ImagePaths.Add(imageCheckBox[i]);
            }
            GeneratePowerPointFile(model);

            return View("Index");
        }

        //This method creates a powerpoint file using the Title and Message text along with the images that were passed in
        public void GeneratePowerPointFile(RichTextEditorViewModel model)
        {
            try
            {
                string titleText = model.Title;
                if (titleText != null)
                {
                    titleText = titleText.Replace("<p>", "");
                    titleText = titleText.Replace("</p>", "");
                }

                string bodyText = model.Message;
                if (bodyText != null)
                {
                    bodyText = bodyText.Replace("<p>", "");
                    bodyText = bodyText.Replace("</p>", "");
                    bodyText = bodyText.Replace("<b>", "");
                    bodyText = bodyText.Replace("</b>", "");
                }

                Application pptApplication = new Application();

                Microsoft.Office.Interop.PowerPoint.Slides slides;
                Microsoft.Office.Interop.PowerPoint._Slide slide;
                Microsoft.Office.Interop.PowerPoint.TextRange objText;

                // Create the Presentation File
                Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);

                Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];

                // Create new Slide
                slides = pptPresentation.Slides;
                slide = slides.AddSlide(1, customLayout);

                // Add title and body text
                objText = slide.Shapes[1].TextFrame.TextRange;
                objText.Text = titleText;
                objText.Font.Name = "Arial";
                objText.Font.Size = 32;
                objText = slide.Shapes[2].TextFrame.TextRange;
                if (bodyText == null)
                {
                    objText.Text = " ";
                } else 
                { 
                    objText.Text = bodyText;
                }
            Microsoft.Office.Interop.PowerPoint.Shape shape = slide.Shapes[2];

                if (model.ImagePaths != null && model.ImagePaths.Count() > 0)
                {
                    slide.Shapes.AddPicture(model.ImagePaths[0], Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 66, shape.Top + 35, 400, shape.Height);
                }
                if (model.ImagePaths != null && model.ImagePaths.Count() > 1)
                {
                    slide.Shapes.AddPicture(model.ImagePaths[1], Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 466, shape.Top + 35, 400, shape.Height);
                }



                slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "Leo Ramos - 2020";



                pptPresentation.SaveAs(@"c:\temp\fppt.pptx", Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                //pptPresentation.Close();
                //pptApplication.Quit();
            }
            catch
            {
                ErrorPage();
            }


        }

        public ActionResult ErrorPage()
        {
            return View();
        }

    }
}