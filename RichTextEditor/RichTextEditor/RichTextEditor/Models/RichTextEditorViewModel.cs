using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using Microsoft.AspNet.Identity;
using Microsoft.Owin.Security;
using System.Web.Mvc;

namespace RichTextEditor.Models
{
    public class RichTextEditorViewModel
    {
        [AllowHtml]
        [Display(Name = "Title")]
        public string Title { get; set; }

        [AllowHtml]
        [Display(Name = "Message")]
        public string Message { get; set; }
        public List<string> ImagePaths { get; set; }
    }
}