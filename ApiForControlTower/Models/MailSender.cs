namespace ApiForControlTower.Controllers
{
    public partial class MailController
    {
        public class MailSender
        {
            public string Subject { get; set; }
            public string From { get; set; }
            public string[] ToList { get; set; }
            public string[] CCList { get; set; }
            public string[] FilePathList { get; set; }
            public string Content { get; set; }

        }
    }
}
