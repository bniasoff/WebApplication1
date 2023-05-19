
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Domain.Models
{
    public class CalendarEvent
    {
        [Key]
        public string EventID { get; set; } 
        public string? FamilyName { get; set; }
        //public string CalendarName { get; set; }
        //public string EventName { get; set; }
        public string? Location { get; set; }
        public string? Address { get; set; }
        public DateTime EventDate { get; set; }
        public TimeSpan? StartTime { get; set; }
        public TimeSpan? EndTime { get; set; }
        public string? Phone { get; set; }
        public string? Category { get; set; }
        [Column(TypeName = "money")]
        public decimal? Charge { get; set; }
        [Column(TypeName = "money")]
        public decimal? Paid { get; set; }
        public bool? ToDo { get; set; }
        public bool? Ready { get; set; }
        public bool? Sent { get; set; }
        public string? Referred { get; set; }
        public DateTime CreatedDate { get; set; }
        public  string ETagID { get; set; }
   
    }      

}





 