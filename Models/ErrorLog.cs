using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Tools.Models
{
    public class ErrorLog
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int ErrorId { get; set; }
        public string Error { get; set; }
        public string Message { get; set; }
        public string Occurance {  get; set; }
        public DateTime LoggedAt { get; set; }
    }
}
