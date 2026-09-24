using ServerExcelRabota.AppDbContext; 
using ServerExcelRabota.Model; 
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace ServerExcelRabota.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class LicenseController : ControllerBase
    {
        private readonly AppDbContextCs _context;

        public LicenseController(AppDbContextCs context)
        {
            _context = context;
        }

        public class ActivateRequest
        {
            public string Key { get; set; }
            public string Hwid { get; set; }
        }

        [HttpPost("activate")]
        public async Task<IActionResult> Activate([FromBody] ActivateRequest request)
        {
            if (string.IsNullOrEmpty(request.Key) || string.IsNullOrEmpty(request.Hwid))
                return BadRequest(new { message = "Ключ или HWID не переданы" });

            var license = await _context.Licenses.FirstOrDefaultAsync(l=>l.LicenseKey==request.Key);

            if (license == null)
                return NotFound(new { message = "Лицензионный ключ не найден" });

            if (!license.IsActive)
                return StatusCode(403, new { message = "Лицензия заблокирована" });

            if (!string.IsNullOrEmpty(license.HardwareId) && license.HardwareId != request.Hwid)
                return StatusCode(403, new { message = "Ключ уже активирован на другом компьютере" });

            if (license.HardwareId == request.Hwid)
                return Ok(new { message = "Лицензия уже активирована на этом устройстве" });


            license.HardwareId = request.Hwid;
            license.ActivatedAt = DateTime.UtcNow;

            await _context.SaveChangesAsync();

            return Ok(new { message = "Лицензия успешно активирована!" });
        }

        [HttpPost("check")]
        public async Task<IActionResult> Check([FromBody] ActivateRequest request)
        {
            if (string.IsNullOrEmpty(request.Key) || string.IsNullOrEmpty(request.Hwid))
                return BadRequest(new { message = "Неверные данные" });

            var license = await _context.Licenses.FirstOrDefaultAsync(l => l.LicenseKey == request.Key);

            if (license == null)
                return NotFound(new { message = "Ключ не найден" });

            if (!license.IsActive)
                return StatusCode(403, new { message = "Лицензия отключена" });

            if (license.HardwareId != request.Hwid)
                return StatusCode(403, new { message = "Лицензия не привязана к этому компьютеру" });



            return Ok(new { message = "OK" });
        }


        [HttpPatch("reset/{key}")]
        public async Task<IActionResult> ResetBinding(string key)
        {
            var license = await _context.Licenses.FirstOrDefaultAsync(l => l.LicenseKey == key);
            if (license == null) return NotFound();

            license.HardwareId = null;
            await _context.SaveChangesAsync();

            return Ok(new { message = "Привязка сброшена. Клиент может активировать ключ заново." });
        }


        [HttpPost("generate")]
        public async Task<IActionResult> GenerateKey()
        {
            string newKey = Guid.NewGuid().ToString("N").Substring(0, 16).ToUpper(); 

            var license = new License
            {
                LicenseKey = newKey,
                IsActive = true,
                CreatedAt = DateTime.UtcNow
            };

            _context.Licenses.Add(license);
            await _context.SaveChangesAsync();

            return Ok(new { key = newKey });
        }
    }
}