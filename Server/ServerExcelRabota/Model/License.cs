using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ServerExcelRabota.Model
{
    [Table("licenses")]
    public class License
    {
        [Key]
        [Column("id")]
        public int Id { get; set; }

        [Required]
        [Column("license_key")]
        public string LicenseKey { get; set; }

        [Column("hardware_id")]
        public string? HardwareId { get; set; } 

        [Column("is_active")]
        public bool IsActive { get; set; } = true;

        [Column("created_at")]
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        [Column("activated_at")]
        public DateTime? ActivatedAt { get; set; }
    }
   
}
