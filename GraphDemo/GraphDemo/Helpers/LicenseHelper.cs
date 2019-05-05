using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace GraphDemo
{
    public class LicenseHelper
    {
        private GraphServiceClient _graphClient;

        public LicenseHelper(GraphServiceClient graphClient)
        {
            if (null == graphClient) throw new ArgumentNullException(nameof(graphClient));
            _graphClient = graphClient;
        }

        public async Task<SubscribedSku> GetLicense()
        {
            var skuResult = await _graphClient.SubscribedSkus.Request().GetAsync();
            return skuResult[1];
        }

        public async Task AddLicense(string userId, Guid? skuId)
        {
            var licensesToAdd = new List<AssignedLicense>();
            var licensesToRemove = new List<Guid>();

            var license = new AssignedLicense()
            {
                SkuId = skuId,
            };

            licensesToAdd.Add(license);

            await _graphClient.Users[userId].AssignLicense(licensesToAdd, licensesToRemove).Request().PostAsync();
        }
    }
}
