using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;

namespace HelloWorld.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class GraphController : ControllerBase
    {
        private readonly GraphAuthenticator _graphAuthenticator;
        private readonly GraphService _graphService;
        public GraphController(GraphAuthenticator authenticator, GraphService service)
        {
            _graphAuthenticator = authenticator;
            _graphService = service;

            _graphService.Authenticate(_graphAuthenticator);
        }

        [HttpGet]
        [Route("{*url}")]
        public async Task<IActionResult> GetAsync(string url)
        {
            string responseData;
            try
            {
                var SSOToken = this.HttpContext.Request.Headers["Authorization"];

                var response = await _graphService.ProcessRequestAsync("GET", url, null, SSOToken, Request.ContentType);
                responseData = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    var data = System.Text.Json.JsonSerializer.Deserialize<object>(responseData);

                    return Ok(new
                    {
                        Data = data
                    });
                }
            }
            catch (System.Exception ex)
            {

                responseData = ex.Message;
            }


            return Problem(responseData, statusCode: 500);

        }
    }
}
