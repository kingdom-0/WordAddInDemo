﻿using System;

namespace WordAddInDemoV2.Helpers
{
    internal class GuidGenerator
    {
        public static string NewGuid()
        {
            return Guid.NewGuid().ToString();
        }
    }
}
