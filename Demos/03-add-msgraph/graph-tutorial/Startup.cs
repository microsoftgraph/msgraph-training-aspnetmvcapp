// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.
using Microsoft.Owin;
using Owin;

[assembly: OwinStartup(typeof(graph_tutorial.Startup))]

namespace graph_tutorial
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
