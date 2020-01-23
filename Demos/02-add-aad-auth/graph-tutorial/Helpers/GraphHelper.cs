// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace graph_tutorial.Helpers
{
    public static class GraphHelper
    {
        public static async Task<User> GetUserDetailsAsync(string accessToken)
        {
            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", accessToken);
                        return Task.FromResult(0);
                    }));

            return await graphClient.Me.Request().GetAsync();
        }
    }
}