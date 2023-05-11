using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoAppenWinform.Http
{
    public class HTTPHelper
    {

        public async Task<string> GetProxyUrlAsync()
        {
            var httpClient = new HttpClient();
            var response = await httpClient.GetAsync("http://10.20.2.192:9049/v1/ips?num=1&country=us&state=WA");
            var responseContent = await response.Content.ReadAsStringAsync();
            return responseContent;
        }
    }
}
