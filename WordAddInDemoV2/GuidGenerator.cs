using System;

namespace WordAddInDemoV2
{
    internal class GuidGenerator
    {
        public static string NewGuid()
        {
            return Guid.NewGuid().ToString();
        }
    }
}
